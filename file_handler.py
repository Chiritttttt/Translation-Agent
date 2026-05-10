import os
import base64
import sys

# ─── OCR 进度回调 ────────────────────────────────────
_ocr_progress_callback = None


def set_ocr_progress_callback(callback):
    """设置 OCR 进度回调函数，供 GUI 调用。
    callback(current_page, total_pages, message)
    """
    global _ocr_progress_callback
    _ocr_progress_callback = callback


def _report_ocr_progress(current, total, msg):
    if _ocr_progress_callback:
        try:
            _ocr_progress_callback(current, total, msg)
        except Exception:
            pass


# ─── 读取 ────────────────────────────────────────────────

def read_pdf(file_path):
    """读取 PDF 文件，自动判断是否为扫描件并启用 AI OCR。

    流程：
    1. 用 PyMuPDF 提取文本
    2. 质量检测：平均每页 < 20 字符 → 判定为扫描件
    3. 扫描件 → pdf2image 转图片 → OpenAI 兼容 Vision API 做 OCR
    """
    try:
        import fitz
        doc = fitz.open(file_path)
        page_count = len(doc)

        # 第一步：尝试文本提取
        text_parts = []
        for page in doc:
            page_text = page.get_text().strip()
            text_parts.append(page_text)
        full_text = "\n\n".join(text_parts).strip()

        # 第二步：质量检测
        avg_chars = len(full_text) / max(page_count, 1)
        if avg_chars >= 20:
            return full_text

        # 第三步：文本太少，很可能是扫描件，启用 OCR
        _report_ocr_progress(0, page_count, "检测到扫描件，启用 AI 视觉 OCR...")
        ocr_text = _ocr_with_vision(file_path, page_count)
        if ocr_text and len(ocr_text.strip()) > 10:
            return ocr_text.strip()

        # OCR 也失败，返回原始文本（可能为空）
        return full_text

    except Exception as e:
        raise Exception(f"PDF读取失败: {e}")


def _ocr_with_vision(file_path, page_count):
    """使用 OpenAI 兼容的 Vision API 对 PDF 逐页做 OCR。
    读取 .env 中的 OPENAI_API_KEY / OPENAI_BASE_URL / OPENAI_MODEL 配置。
    """
    try:
        from dotenv import load_dotenv
        load_dotenv()

        import openai
        from pdf2image import convert_from_path
        from io import BytesIO

        api_key = os.getenv("OPENAI_API_KEY", "")
        base_url = os.getenv("OPENAI_BASE_URL", "")
        model = os.getenv("OPENAI_MODEL", "")
        # OCR 可能需要不同的 Vision 模型（DeepSeek 文本模型不支持图像输入）
        vision_model = os.getenv("OPENAI_VISION_MODEL", "") or model

        if not api_key:
            print("[OCR] 未配置 OPENAI_API_KEY，无法进行 OCR 识别")
            return ""

        # 选择模型：优先使用用户配置的 Vision 模型
        if not vision_model:
            vision_model = "gpt-4o"

        client = openai.OpenAI(api_key=api_key, base_url=base_url if base_url else None)

        # 将 PDF 每页转为图片
        _report_ocr_progress(0, page_count, "正在转换 PDF 页面为图片...")
        images = convert_from_path(file_path, dpi=200)

        if not images:
            return ""

        actual_pages = len(images)
        all_text = []

        for i, img in enumerate(images):
            _report_ocr_progress(i + 1, actual_pages,
                                 f"OCR 识别中: 第 {i+1}/{actual_pages} 页...")
            try:
                # 将 PIL Image 转 base64
                buf = BytesIO()
                # 统一转为 RGB 避免 RGBA 问题
                if img.mode == "RGBA":
                    img = img.convert("RGB")
                img.save(buf, format="PNG")
                img_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

                # 调用 Vision API
                response = client.chat.completions.create(
                    model=vision_model,
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "你是一个专业的 OCR 助手。请从文档图片中完整、准确地提取所有文字。"
                                "保留段落、标题、列表等结构。只输出提取的文字，不要有任何解释。"
                            ),
                        },
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "text",
                                    "text": f"请提取第 {i+1} 页的所有文字，保留原始格式和结构。",
                                },
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:image/png;base64,{img_b64}"
                                    },
                                },
                            ],
                        },
                    ],
                    temperature=0.1,
                    max_tokens=16384,
                )

                page_text = response.choices[0].message.content or ""
                if page_text.strip():
                    all_text.append(page_text.strip())

            except Exception as e:
                print(f"[OCR] 第 {i+1} 页识别失败: {e}")
                continue

        _report_ocr_progress(actual_pages, actual_pages, "OCR 识别完成！")
        return "\n\n".join(all_text)

    except ImportError as e:
        print(f"[OCR] 缺少依赖: {e}")
        print("[OCR] 请安装: pip install pdf2image Pillow openai python-dotenv")
        return ""
    except Exception as e:
        print(f"[OCR] Vision API 调用失败: {e}")
        return ""


def read_docx(file_path):
    try:
        from docx import Document
        doc = Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs])
        return text.strip()
    except Exception as e:
        raise Exception(f"DOCX读取失败: {e}")


def read_pptx(file_path):
    """读取 PPTX，按幻灯片逐页提取文本，保留结构标记。

    输出格式：
        [翻译说明] 以下为PPT幻灯片逐页内容...
        [幻灯片 1/N]
        标题文本
        正文行1
        正文行2
        [幻灯片 2/N]
        ...

    支持识别：标题、正文、表格、组合形状、SmartArt。
    """
    try:
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE

        prs = Presentation(file_path)
        total = len(prs.slides)
        parts = []

        # AI 翻译说明行，告诉 AI 保留标记
        parts.append(
            "[翻译说明] 以下为PPT幻灯片逐页内容。"
            "请保留所有 [幻灯片 N/M] 标记，只翻译标记后的文字。"
            "标题保持简洁有力，正文翻译准确流畅。"
        )
        parts.append("")

        for idx, slide in enumerate(prs.slides, 1):
            title_text = ""
            body_lines = []

            for shape in slide.shapes:
                # ── 组合形状 / SmartArt ──
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    body_lines.extend(_extract_group_text(shape))
                    continue

                # ── 表格 ──
                if shape.has_table:
                    body_lines.extend(_extract_table_text(shape.table))
                    continue

                # ── 文本框 ──
                if shape.has_text_frame:
                    shape_text = shape.text.strip()
                    if not shape_text:
                        continue

                    if _is_title_shape(shape):
                        title_text = shape_text
                    else:
                        for para in shape.text_frame.paragraphs:
                            t = para.text.strip()
                            if t:
                                body_lines.append(t)

            # 为每张幻灯片输出标记（包括纯图片页，保持编号一致）
            parts.append(f"[幻灯片 {idx}/{total}]")
            if title_text:
                parts.append(title_text)
            if body_lines:
                parts.extend(body_lines)
            parts.append("")  # 空行分隔各页

        return "\n".join(parts).strip()

    except Exception as e:
        raise Exception(f"PPTX读取失败: {e}")


def read_excel(file_path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True)
        text = ""
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                row_text = "  ".join([str(c) for c in row if c is not None])
                if row_text.strip():
                    text += row_text + "\n"
        return text.strip()
    except Exception as e:
        raise Exception(f"Excel读取失败: {e}")


def read_file(file_path):
    """读取文件，自动识别格式。

    返回: (file_type, content, extra)
      - file_type: 'pdf' | 'docx' | 'pptx' | 'xlsx' | 'txt'
      - content: 提取的文本
      - extra: Excel 时为行数据列表 [[cell, ...], ...]，其他格式为 None
    """
    ext = file_path.lower().split(".")[-1]
    if ext == "pdf":
        return "pdf", read_pdf(file_path), None
    elif ext in ["docx", "doc"]:
        return "docx", read_docx(file_path), None
    elif ext in ["pptx", "ppt"]:
        return "pptx", read_pptx(file_path), None
    elif ext in ["xlsx", "xls"]:
        return "xlsx", read_excel(file_path), None
    else:
        raise Exception("不支持此格式")


# ─── 工具：译文行对齐 ────────────────────────────────────

def match_translation(original_lines, translated_text):
    """将译文行与原文行对齐。

    策略（按优先级）：
    1. 行数相同 → 直接按序 1:1 匹配
    2. 行数不同 → 用 LCS（最长公共子序列）对齐，保留顺序
    3. 回退 → 贪心左对齐，每条原文消耗一条译文
    """
    import difflib

    trans_lines = [l for l in translated_text.split("\n") if l.strip()]
    orig_count = len(original_lines)
    trans_count = len(trans_lines)
    if orig_count == 0:
        return []
    if trans_count == 0:
        return [""] * orig_count

    # ── 路径 A：行数完全相同，直接按序匹配 ──
    if orig_count == trans_count:
        return list(trans_lines[:orig_count])

    # ── 路径 B：行数不同，用 LCS 对齐 ──
    # 将原文和译文视为两个序列，找出最长公共子序列，
    # 然后将非公共部分（新增/删除的行）均匀分配到相邻的公共行之间。
    # 这样即使 AI 合并或拆分了某些行，也能正确对齐。
    result = _lcs_align(original_lines, trans_lines)
    if result is not None:
        return result

    # ── 路径 C：回退到贪心左对齐 ──
    result = []
    trans_idx = 0
    for i in range(orig_count):
        if trans_idx < trans_count:
            result.append(trans_lines[trans_idx])
            trans_idx += 1
        else:
            result.append(original_lines[i])
    return result


def _lcs_align(orig_lines, trans_lines):
    """基于 LCS 的智能对齐。

    通过 difflib.SequenceMatcher 找出原文和译文中"对应"的行
    （基于文本相似度），然后将差异部分均匀分配。
    返回与 orig_lines 等长的列表，每项是对应的译文。
    """
    import difflib

    # 构建相似度矩阵：用文本长度比和字符重叠度判断哪些译文行
    # "对应"哪些原文行
    orig_count = len(orig_lines)
    trans_count = len(trans_lines)

    # 快速路径：行数差异太大时 LCS 对齐效果差，回退到贪心
    ratio = max(orig_count, trans_count) / (min(orig_count, trans_count) + 1)
    if ratio > 5:
        return None

    # 用 SequenceMatcher 找出两个序列的最长公共子序列
    # 但我们需要的是"语义对齐"而非"文本完全匹配"
    # 所以用一个简化方案：将译文按比例切分分配给原文段落
    #
    # 更实用的方案：计算每段原文的大致字符占比，
    # 按比例将译文行分配到对应的原文段落
    orig_char_counts = [len(l) for l in orig_lines]
    total_orig_chars = sum(orig_char_counts)
    if total_orig_chars == 0:
        return None

    trans_char_counts = [len(l) for l in trans_lines]
    total_trans_chars = sum(trans_char_counts)
    if total_trans_chars == 0:
        return None

    # 按字符比例分配：每段原文应分到的字符数
    result = []
    trans_char_budget = 0.0
    trans_idx = 0
    accumulated_trans_text = ""

    for i in range(orig_count):
        # 这段原文应分到的译文字符数
        proportion = orig_char_counts[i] / total_orig_chars
        target_chars = proportion * total_trans_chars
        trans_char_budget += target_chars

        # 消耗译文行直到达到预算
        while trans_idx < trans_count:
            accumulated_trans_text += trans_lines[trans_idx]
            trans_idx += 1
            if len(accumulated_trans_text) >= trans_char_budget * 0.8:
                break

        result.append(accumulated_trans_text.strip())
        accumulated_trans_text = ""

    # 处理剩余译文行（归入最后一段）
    remaining = [trans_lines[j] for j in range(trans_idx, trans_count)]
    if remaining:
        last_text = "\n".join(remaining)
        if result:
            result[-1] = (result[-1] + "\n" + last_text).strip()
        else:
            result.append(last_text)

    # 确保返回与 orig_lines 等长的列表
    while len(result) < orig_count:
        result.append(orig_lines[len(result)])

    return result[:orig_count]


# ─── 工具：复制 DOCX 格式 ────────────────────────────────

def _copy_run_format(src_run, dst_run):
    try:
        dst_run.bold = src_run.bold
        dst_run.italic = src_run.italic
        dst_run.underline = src_run.underline
        if src_run.font.size:
            dst_run.font.size = src_run.font.size
        if src_run.font.color and src_run.font.color.type:
            try:
                dst_run.font.color.rgb = src_run.font.color.rgb
            except Exception:
                pass
        if src_run.font.name:
            dst_run.font.name = src_run.font.name
    except Exception:
        pass


def _copy_para_format(src_para, dst_para):
    try:
        dst_para.alignment = src_para.alignment
        try:
            dst_para.style = src_para.style
        except Exception:
            pass
        pf = src_para.paragraph_format
        dpf = dst_para.paragraph_format
        if pf.space_before:
            dpf.space_before = pf.space_before
        if pf.space_after:
            dpf.space_after = pf.space_after
        if pf.line_spacing:
            dpf.line_spacing = pf.line_spacing
        if pf.left_indent:
            dpf.left_indent = pf.left_indent
        if pf.first_line_indent:
            dpf.first_line_indent = pf.first_line_indent
    except Exception:
        pass


# ─── 导出 DOCX ───────────────────────────────────────────

def export_docx_translation(translated, output_path, original_path=None):
    from docx import Document
    if original_path and os.path.exists(original_path):
        src_doc = Document(original_path)
        dst_doc = Document(original_path)
        orig_paras = [p for p in src_doc.paragraphs if p.text.strip()]
        matched = match_translation([p.text for p in orig_paras], translated)
        trans_idx = 0
        for para in dst_doc.paragraphs:
            if para.text.strip() and trans_idx < len(matched):
                new_text = matched[trans_idx]
                trans_idx += 1
                if para.runs:
                    para.runs[0].text = new_text
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.add_run(new_text)
        dst_doc.save(output_path)
    else:
        doc = Document()
        for line in translated.split("\n"):
            if line.strip():
                doc.add_paragraph(line)
        doc.save(output_path)


def export_docx_paragraph(original, translated, output_path, original_path=None):
    """段落对照：一段原文一段译文，完全复制原文格式，无装饰"""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    import copy

    if original_path and os.path.exists(original_path):
        src_doc = Document(original_path)
        orig_src_paras = [p for p in src_doc.paragraphs if p.text.strip()]
    else:
        orig_src_paras = None

    # 从原文文件直接复制作为基础
    if original_path and os.path.exists(original_path):
        dst_doc = Document(original_path)
        # 清空所有段落内容
        for para in dst_doc.paragraphs:
            for run in para.runs:
                run.text = ""
    else:
        dst_doc = Document()

    raw_orig = [p for p in original.split("\n") if p.strip()]
    matched_trans = match_translation(raw_orig, translated)

    # 清除文档内容，重新写入
    from docx.oxml.ns import qn
    body = dst_doc.element.body
    # 保留最后一个sectPr（页面设置），清除其余内容
    sect_pr = body.find(qn('w:sectPr'))
    for child in list(body):
        if child != sect_pr:
            body.remove(child)

    for i, orig_text in enumerate(raw_orig):
        # ── 原文段落：完整复制格式 ──
        orig_p = dst_doc.add_paragraph()
        if orig_src_paras and i < len(orig_src_paras):
            _copy_para_format(orig_src_paras[i], orig_p)
            for src_run in orig_src_paras[i].runs:
                dst_run = orig_p.add_run(src_run.text)
                _copy_run_format(src_run, dst_run)
        else:
            orig_p.add_run(orig_text)

        # ── 译文段落：完全复制原文格式，只换文字 ──
        trans_text = matched_trans[i] if i < len(matched_trans) else ""
        trans_p = dst_doc.add_paragraph()
        if orig_src_paras and i < len(orig_src_paras):
            _copy_para_format(orig_src_paras[i], trans_p)
            if orig_src_paras[i].runs:
                # 只用第一个run的格式，文字换成译文
                dst_run = trans_p.add_run(trans_text)
                _copy_run_format(orig_src_paras[i].runs[0], dst_run)
            else:
                trans_p.add_run(trans_text)
        else:
            trans_p.add_run(trans_text)

    dst_doc.save(output_path)


def export_docx_bilingual(original, translated, output_path, original_path=None):
    """行对照：双列表格，去掉标题"""
    from docx import Document
    from docx.shared import RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT

    if original_path and os.path.exists(original_path):
        src_doc = Document(original_path)
        orig_src_paras = [p for p in src_doc.paragraphs if p.text.strip()]
    else:
        orig_src_paras = None

    doc = Document()
    # ← 去掉了 add_heading
    raw_orig = [p for p in original.split("\n") if p.strip()]
    matched_trans = match_translation(raw_orig, translated)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for idx, txt in enumerate(["原文", "译文"]):
        run = hdr[idx].paragraphs[0].add_run(txt)
        run.bold = True
        run.font.color.rgb = RGBColor(0x4F, 0x8E, 0xF7)

    for i, orig_text in enumerate(raw_orig):
        row = table.add_row().cells
        orig_cell_para = row[0].paragraphs[0]
        if orig_src_paras and i < len(orig_src_paras):
            _copy_para_format(orig_src_paras[i], orig_cell_para)
            for src_run in orig_src_paras[i].runs:
                dst_run = orig_cell_para.add_run(src_run.text)
                _copy_run_format(src_run, dst_run)
        else:
            orig_cell_para.add_run(orig_text)

        trans_cell_para = row[1].paragraphs[0]
        trans_text = matched_trans[i] if i < len(matched_trans) else ""
        if orig_src_paras and i < len(orig_src_paras):
            _copy_para_format(orig_src_paras[i], trans_cell_para)
            if orig_src_paras[i].runs:
                dst_run = trans_cell_para.add_run(trans_text)
                _copy_run_format(orig_src_paras[i].runs[0], dst_run)
            else:
                trans_cell_para.add_run(trans_text)
        else:
            trans_cell_para.add_run(trans_text)

    doc.save(output_path)


# ═══════════════════════════════════════════════════════════
# PPT 辅助函数
# ═══════════════════════════════════════════════════════════

def _is_title_shape(shape):
    """判断 shape 是否为标题占位符"""
    try:
        if not hasattr(shape, "placeholder_format") or shape.placeholder_format is None:
            return False
        from pptx.enum.shapes import PP_PLACEHOLDER
        ph_type = shape.placeholder_format.type
        # TITLE (1), SUBTITLE (2), CENTER_TITLE (3), TITLE_2 (7)
        return ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.SUBTITLE,
                           PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE_2)
    except Exception:
        return False


def _extract_group_text(group_shape):
    """递归提取组合形状 / SmartArt 中的文本"""
    lines = []
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        for child in group_shape.shapes:
            if child.shape_type == MSO_SHAPE_TYPE.GROUP:
                lines.extend(_extract_group_text(child))
            elif child.has_text_frame:
                for para in child.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        lines.append(t)
            elif child.has_table:
                lines.extend(_extract_table_text(child.table))
    except Exception:
        pass
    return lines


def _extract_table_text(table):
    """提取 PPT 表格中的文本，按行列输出"""
    lines = []
    try:
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                t = cell.text.strip()
                row_cells.append(t if t else "")
            lines.append(" | ".join(row_cells))
    except Exception:
        pass
    return lines


def _collect_slide_paragraphs(slide):
    """收集单张幻灯片中所有有文本的段落及其所属 shape 信息。

    表格处理：与 read_pptx 的 _extract_table_text 保持一致，
    按行输出 "cell1 | cell2 | cell3" 格式，而不是逐单元格。
    每行对应一个条目，para 指向该行第一个单元格的第一个段落。

    Returns:
        list of dict: [
            {"shape": shape, "para": para, "is_title": bool, "text": str,
             "is_table_row": bool, "row_cells": [para, ...]},
            ...
        ]
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    result = []

    for shape in slide.shapes:
        # 组合形状：递归收集
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for item in _collect_group_paragraphs(shape):
                result.append(item)
            continue

        # 表格：按行收集（与 read_pptx _extract_table_text 一致）
        if shape.has_table:
            for row in shape.table.rows:
                row_paras = []
                row_texts = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    # 取每个单元格的第一个有文本的段落
                    first_para = None
                    for para in cell.text_frame.paragraphs:
                        if para.text.strip():
                            if first_para is None:
                                first_para = para
                            break
                    row_paras.append(first_para)
                    row_texts.append(cell_text if cell_text else "")

                row_text = " | ".join(row_texts)
                if row_text.strip():
                    result.append({
                        "shape": shape,
                        "para": row_paras[0],  # 该行第一个有内容的段落
                        "is_title": False,
                        "text": row_text,
                        "is_table_row": True,
                        "row_cells": row_paras,
                    })
            continue

        # 文本框
        if shape.has_text_frame:
            is_title = _is_title_shape(shape)
            for para in shape.text_frame.paragraphs:
                if para.text.strip():
                    result.append({
                        "shape": shape,
                        "para": para,
                        "is_title": is_title,
                        "text": para.text.strip(),
                        "is_table_row": False,
                        "row_cells": None,
                    })

    return result


def _collect_group_paragraphs(group_shape):
    """递归收集组合形状中的段落"""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    result = []
    try:
        for child in group_shape.shapes:
            if child.shape_type == MSO_SHAPE_TYPE.GROUP:
                result.extend(_collect_group_paragraphs(child))
            elif child.has_text_frame:
                for para in child.text_frame.paragraphs:
                    if para.text.strip():
                        result.append({
                            "shape": child,
                            "para": para,
                            "is_title": False,
                            "text": para.text.strip(),
                            "is_table_cell": False,
                        })
    except Exception:
        pass
    return result


def _reconstruct_slide_markers(original_path, translated_text):
    """如果 AI 译文丢失了 [幻灯片 N/M] 标记，根据原始 PPT 结构自动重建。

    这是最后一道防线：即使 AI 在翻译/审校/终稿步骤中丢掉了标记，
    导出时也能根据原始 PPT 的页面结构，按比例将译文分配到每一页。

    Returns:
        str: 带有 [幻灯片 N/M] 标记的译文文本
    """
    import re

    # 已有标记，无需处理
    if re.search(r'\[幻灯片\s+\d+/\d+\]', translated_text):
        return translated_text

    # 读取原始 PPT 的文本结构
    try:
        source_text = read_pptx(original_path)
    except Exception:
        return translated_text

    # 解析原文的幻灯片结构：每页有多少行内容
    structure = []  # [(slide_num, total_slides, content_line_count)]
    lines = source_text.split("\n")
    current_slide = None
    total_slides = 0
    content_count = 0

    for line in lines:
        m = re.match(r'^\s*\[幻灯片\s+(\d+)/(\d+)\]', line)
        if m:
            if current_slide is not None:
                structure.append((current_slide, total_slides, content_count))
            current_slide = int(m.group(1))
            total_slides = int(m.group(2))
            content_count = 0
        elif line.strip() and not line.startswith("[翻译说明]"):
            content_count += 1

    if current_slide is not None:
        structure.append((current_slide, total_slides, content_count))

    if not structure:
        return translated_text

    # 统计原文总内容行数
    total_content = sum(s[2] for s in structure)
    if total_content == 0:
        # 所有页面都是空白的，均匀分配
        total_content = len(structure)

    # 提取译文的非空行
    trans_lines = [l.strip() for l in translated_text.split("\n")
                   if l.strip()
                   and not l.startswith("[翻译说明]")
                   and not re.match(r'^\s*\[幻灯片', l)]

    if not trans_lines:
        return translated_text

    # 按比例将译文行分配到每一页
    result_parts = []
    trans_idx = 0

    for i, (slide_num, total, orig_count) in enumerate(structure):
        # 计算这一页应分配多少行译文
        if orig_count == 0:
            expected = max(1, len(trans_lines) // len(structure))
        else:
            proportion = orig_count / total_content
            expected = max(1, round(proportion * len(trans_lines)))

        # 最后一页：把剩余行全部给最后一页
        if i == len(structure) - 1:
            expected = len(trans_lines) - trans_idx

        # 取译文行
        end_idx = min(trans_idx + expected, len(trans_lines))
        slide_content = trans_lines[trans_idx:end_idx]

        result_parts.append(f"[幻灯片 {slide_num}/{total}]")
        result_parts.extend(slide_content)
        result_parts.append("")

        trans_idx = end_idx

    return "\n".join(result_parts)


def _parse_translated_slides(text):
    """将 AI 译文按 [幻灯片 N/M] 标记解析为逐页内容。

    Returns:
        dict: {slide_number: [translated_lines]}
        如果没有标记，返回 {1: [all_lines]}（兼容旧格式）
    """
    import re
    slides = {}
    current_slide = None
    current_lines = []

    for line in text.split("\n"):
        # 跳过翻译说明行
        if line.startswith("[翻译说明]") or line.startswith("[翻译说明]"):
            continue

        # 匹配幻灯片标记（兼容多种空格写法）
        m = re.match(r'^\s*\[幻灯片\s+(\d+)', line)
        if m:
            # 保存上一页
            if current_slide is not None:
                slides[current_slide] = [l for l in current_lines if l.strip()]
            current_slide = int(m.group(1))
            current_lines = []
            continue

        if current_slide is not None:
            current_lines.append(line.strip())

    # 最后一页
    if current_slide is not None:
        slides[current_slide] = [l for l in current_lines if l.strip()]

    # 兼容：没有标记的纯文本，全部作为第 1 页
    if not slides:
        all_lines = [l.strip() for l in text.split("\n") if l.strip()
                     and not l.startswith("[翻译说明]")]
        if all_lines:
            slides[1] = all_lines

    return slides


def _replace_para_text(para, new_text):
    """替换段落文本，保留第一个 run 的格式。"""
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.text = new_text


def _split_title_and_body(lines):
    """从译文行列表中分离标题和正文。

    规则：第一行短文本（≤80字符）视为标题，其余为正文。
    如果只有一行，不区分。
    注意：阈值从60放宽到80，因为中文翻译往往比英文原文更长。
    """
    if not lines:
        return "", []
    if len(lines) == 1:
        return lines[0], []
    # 第一行较短 → 视为标题（阈值80，兼顾中英文标题长度差异）
    if len(lines[0]) <= 80:
        return lines[0], lines[1:]
    return "", lines


# ═══════════════════════════════════════════════════════════
# PPT 导出函数
# ═══════════════════════════════════════════════════════════

# ── 轻量级导出（zipfile 直接操作，不加载图片到内存） ──

def _modify_slide_xml(xml_bytes, trans_lines):
    """修改幻灯片 XML 中的文本内容，保留原始 XML 结构。

    改进版：
    - 区分标题占位符和正文，分别匹配译文行
    - 检测表格结构（<a:tbl>），将单元格段落按行分组，
      与 read_pptx 的行级输出（"cell1 | cell2 | cell3"）对齐
    - 非表格段落逐个匹配

    read_pptx 提取顺序是「标题 → 正文行（含表格行）」，译文也是同样顺序，
    因此必须分开匹配，否则会导致标题和正文错位。

    原理：
    - PPTX 中文本存储在 <a:t> 标签中
    - 按段落级别匹配原文与译文
    - 将译文写入段落第一个 <a:t>，其余清空
    - 全程操作 XML 字符串，不解析为对象，保留所有格式/命名空间
    """
    import re

    xml_str = xml_bytes.decode("utf-8")

    # 1. 找到所有 <a:t> 元素的位置
    t_pat = re.compile(r"(<a:t(?:\s[^>]*)?>)(.*?)(</a:t>)", re.DOTALL)
    all_t = list(t_pat.finditer(xml_str))

    if not all_t:
        return xml_bytes

    # 2. 按段落 <a:p> 分组
    p_open_pos = [m.start() for m in re.finditer(r"<a:p[\s>]", xml_str)]
    p_close_pos = [m.end() for m in re.finditer(r"</a:p>", xml_str)]

    if len(p_open_pos) != len(p_close_pos):
        # 不匹配时退化：每个 <a:t> 独立替换
        orig_texts = [m.group(2) for m in all_t]
        matched = match_translation(orig_texts, "\n".join(trans_lines))
        result = xml_str
        for i in range(len(all_t) - 1, -1, -1):
            m = all_t[i]
            new_text = _xml_escape(matched[i]) if i < len(matched) else m.group(2)
            result = result[:m.start(2)] + new_text + result[m.end(2):]
        return result.encode("utf-8")

    # 3. 收集每个段落信息，并检测标题和表格归属
    _ph_title_re = re.compile(
        r'<p:ph\b[^>]*type\s*=\s*["\'](?:title|ctrTitle|subTitle)["\']'
        r'|<p:ph\b[^>]*idx\s*=\s*["\']0["\']',
    )

    # 检测表格区域：<a:tbl> ... </a:tbl>
    tbl_regions = []
    for tbl_m in re.finditer(r'<a:tbl\b', xml_str):
        tbl_end = re.search(r'</a:tbl>', xml_str[tbl_m.start():])
        if tbl_end:
            tbl_regions.append((tbl_m.start(), tbl_m.start() + tbl_end.end()))

    def _in_table(pos):
        return any(s <= pos <= e for s, e in tbl_regions)

    # 检测表格行区域：<a:tr> ... </a:tr>
    tr_regions = []
    for tr_m in re.finditer(r'<a:tr\b', xml_str):
        tr_end = re.search(r'</a:tr>', xml_str[tr_m.start():])
        if tr_end:
            tr_regions.append((tr_m.start(), tr_m.start() + tr_end.end()))

    def _get_table_row(pos):
        """返回段落所在的表格行索引（在 tr_regions 中的序号），非表格返回 -1"""
        for i, (s, e) in enumerate(tr_regions):
            if s <= pos <= e:
                return i
        return -1

    paragraphs = []  # [{"text", "runs", "is_title", "in_table", "table_row_idx"}, ...]
    t_idx = 0

    for p_start, p_end in zip(p_open_pos, p_close_pos):
        # 检测标题
        search_start = max(0, p_start - 800)
        pre_text = xml_str[search_start:p_start]
        is_title = bool(_ph_title_re.search(pre_text))

        # 收集此段落中的 <a:t>
        runs = []
        full_text = ""
        while t_idx < len(all_t) and all_t[t_idx].start() < p_end:
            m = all_t[t_idx]
            content = m.group(2)
            runs.append((m.start(2), m.end(2), content))
            full_text += content
            t_idx += 1

        if full_text.strip() and runs:
            paragraphs.append({
                "text": full_text,
                "runs": runs,
                "is_title": is_title,
                "in_table": _in_table(p_start),
                "table_row_idx": _get_table_row(p_start),
            })

    if not paragraphs:
        return xml_bytes

    # 4. 分离标题、表格段落组、非表格正文段落
    title_paras = [p for p in paragraphs if p["is_title"]]
    table_paras = [p for p in paragraphs if p["in_table"] and not p["is_title"]]
    non_table_paras = [p for p in paragraphs if not p["in_table"] and not p["is_title"]]

    # 将表格段落按行分组，每行的段落文本用 " | " 拼接（与 read_pptx 一致）
    table_row_groups = {}  # {row_idx: [para, ...]}
    for p in table_paras:
        row_idx = p["table_row_idx"]
        if row_idx >= 0:
            table_row_groups.setdefault(row_idx, []).append(p)

    # 按行索引排序，构建表格行文本列表（与 read_pptx 输出格式一致）
    sorted_row_indices = sorted(table_row_groups.keys())
    table_rows = []
    for ri in sorted_row_indices:
        row_paras = table_row_groups[ri]
        # 每个单元格的第一个段落文本作为该单元格的文本
        # read_pptx 是 cell.text.strip()，即单元格内所有段落拼接
        # 但 _extract_table_text 是按行输出 "cell1 | cell2"
        # 这里我们直接把该行所有段落文本按序拼接作为一行
        row_text = " | ".join(p["text"] for p in row_paras)
        table_rows.append(row_text)

    trans_title, trans_body_lines = _split_title_and_body(trans_lines)

    # 5. 匹配译文
    title_matched = {}
    if trans_title and title_paras:
        title_matched[title_paras[0]["runs"][0][0]] = trans_title

    body_matched = {}

    # 将非表格正文和表格行合并为统一的匹配列表
    # 匹配时按照 read_pptx 的输出顺序：标题之后，先是非表格正文，再是表格行
    # 但实际上 read_pptx 遍历 shapes 的顺序决定了输出顺序
    # 所以我们需要按 XML 中的出现顺序来构建原文列表
    #
    # 简化方案：将非表格段落和表格行交替出现在 XML 中的顺序作为匹配顺序
    non_table_orig = [p["text"] for p in non_table_paras]

    # 构建有序的"原文单元"列表：非表格段落 + 表格行，按 XML 位置排列
    ordered_units = []  # [(type, data), ...] type="para" or "table_row"
    used_table_rows = set()

    for p in paragraphs:
        if p["is_title"]:
            continue
        if p["in_table"] and p["table_row_idx"] >= 0:
            ri = p["table_row_idx"]
            if ri not in used_table_rows:
                used_table_rows.add(ri)
                ordered_units.append(("table_row", ri))
        elif not p["in_table"]:
            ordered_units.append(("para", p))

    # 构建原文文本列表（用于 match_translation）
    orig_texts_for_match = []
    for unit_type, unit_data in ordered_units:
        if unit_type == "table_row":
            ri = unit_data
            # 找到这个行在 sorted_row_indices 中的位置
            if ri in table_row_groups:
                row_text = " | ".join(p["text"] for p in table_row_groups[ri])
                orig_texts_for_match.append(row_text)
        else:
            orig_texts_for_match.append(unit_data["text"])

    # 用译文匹配
    matched_texts = match_translation(
        orig_texts_for_match,
        "\n".join(trans_body_lines),
    )

    # 将匹配结果映射回段落
    for i, (unit_type, unit_data) in enumerate(ordered_units):
        if i < len(matched_texts) and matched_texts[i]:
            if unit_type == "table_row":
                # 表格行：译文是 "cell1 | cell2 | ..." 格式，需要拆分回各单元格
                ri = unit_data
                row_paras = table_row_groups.get(ri, [])
                translated_row = matched_texts[i]
                # 按 " | " 拆分译文
                cells = translated_row.split(" | ")
                # 将译文分配到各单元格段落（每段对应一个单元格）
                for ci, cp in enumerate(row_paras):
                    if ci < len(cells):
                        body_matched[cp["runs"][0][0]] = cells[ci].strip()
            else:
                body_matched[unit_data["runs"][0][0]] = matched_texts[i]

    # 6. 替换：从后往前，避免位置偏移
    for i in range(len(paragraphs) - 1, -1, -1):
        p = paragraphs[i]
        runs = p["runs"]
        if not runs:
            continue

        first_run_start = runs[0][0]

        new_text = title_matched.get(first_run_start) or body_matched.get(first_run_start)
        if not new_text:
            continue

        escaped = _xml_escape(new_text)

        s, e, _ = runs[0]
        xml_str = xml_str[:s] + escaped + xml_str[e:]

        for j in range(len(runs) - 1, 0, -1):
            rs, re_, _ = runs[j]
            xml_str = xml_str[:rs] + xml_str[re_:]

    return xml_str.encode("utf-8")


def _xml_escape(text):
    """转义 XML 特殊字符。"""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def _get_slide_order_mapping(original_path):
    """从 PPTX 的 presentation.xml 读取幻灯片的实际展示顺序。

    PPTX 中 slide 文件名（如 slide1.xml, slide2.xml）不一定等于展示顺序。
    用户在 PowerPoint 中拖动排序幻灯片后，文件名不变，但 presentation.xml
    中的 <p:sldId> 顺序会改变。

    Returns:
        dict: {presentation_order: filename_number}
        例如 {1: 3, 2: 1, 3: 2} 表示展示第1页对应 slide3.xml

        如果读取失败，返回 None（回退到按文件名顺序）。
    """
    import zipfile
    import re

    try:
        with zipfile.ZipFile(original_path, "r") as zf:
            # 读取 presentation.xml
            pres_xml = zf.read("ppt/presentation.xml").decode("utf-8")
            # 读取 presentation.xml.rels 获取 rId → 文件名映射
            rels_xml = ""
            try:
                rels_xml = zf.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
            except KeyError:
                pass

        # 1. 从 presentation.xml 提取 <p:sldId> 的 rId，按出现顺序即展示顺序
        #    格式: <p:sldId id="256" r:embed="rId2"/>
        sld_id_re = re.compile(r'<p:sldId\b[^>]*r:embed=["\'](rId\d+)["\']', re.DOTALL)
        rids_in_order = [m.group(1) for m in sld_id_re.finditer(pres_xml)]

        if not rids_in_order:
            return None

        # 2. 从 rels 文件构建 rId → 文件名映射
        if not rels_xml:
            return None

        rel_re = re.compile(
            r'<Relationship\s+Id=["\'](rId\d+)["\']\s+[^>]*Target=["\']([^"\']+)["\']',
            re.DOTALL,
        )
        rid_to_target = {}
        for m in rel_re.finditer(rels_xml):
            rid = m.group(1)
            target = m.group(2).lstrip("/")
            rid_to_target[rid] = target

        # 3. 构建展示顺序 → 文件名编号的映射
        slide_re = re.compile(r"slide(\d+)\.xml$")
        order_to_filenum = {}
        for order_idx, rid in enumerate(rids_in_order, 1):
            target = rid_to_target.get(rid, "")
            sm = slide_re.search(target)
            if sm:
                order_to_filenum[order_idx] = int(sm.group(1))

        return order_to_filenum if order_to_filenum else None

    except Exception:
        return None


def _export_pptx_zip(original_path, translated, output_path):
    """轻量级 PPT 导出：使用 zipfile 直接操作 ZIP 结构。

    优势：
    - 不把图片/媒体加载到内存 → 支持超大文件
    - 不用 python-pptx 解析整个文件 → 速度快
    - 直接拷贝图片等二进制文件 → 零修改风险

    仅修改 ppt/slides/slideN.xml 中的 <a:t> 文本节点。
    """
    import zipfile
    import re
    import tempfile
    import shutil

    # 解析译文
    translated = _reconstruct_slide_markers(original_path, translated)
    trans_slides = _parse_translated_slides(translated)

    # 获取正确的幻灯片展示顺序映射
    order_to_filenum = _get_slide_order_mapping(original_path)

    slide_re = re.compile(r"ppt/slides/slide(\d+)\.xml$")

    # 构建展示顺序 → 译文的映射
    # trans_slides 的 key 是 read_pptx 中的展示顺序号
    # order_to_filenum 将展示顺序映射到文件名编号
    filenum_to_trans = {}
    if order_to_filenum:
        for order_idx, file_num in order_to_filenum.items():
            if order_idx in trans_slides:
                filenum_to_trans[file_num] = trans_slides[order_idx]
    else:
        # 回退：直接用展示顺序号当文件名编号
        filenum_to_trans = trans_slides

    with zipfile.ZipFile(original_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # 只处理幻灯片 XML
                sm = slide_re.match(item.filename)
                if sm:
                    slide_filenum = int(sm.group(1))
                    if slide_filenum in filenum_to_trans and filenum_to_trans[slide_filenum]:
                        data = zin.read(item.filename)
                        modified = _modify_slide_xml(
                            data, filenum_to_trans[slide_filenum])
                        zout.writestr(item, modified)
                        continue

                # 其他文件（图片、媒体等）→ 原样拷贝
                if item.file_size > 5 * 1024 * 1024:
                    tmp_path = None
                    try:
                        fd, tmp_path = tempfile.mkstemp(suffix=".pptx_tmp")
                        os.close(fd)
                        with zin.open(item) as src, open(tmp_path, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        # 保留原始 ZIP 元数据（时间戳、权限等）
                        zout.write(
                            tmp_path, item.filename,
                            compress_type=item.compress_type,
                            compresslevel=None)
                    finally:
                        if tmp_path and os.path.exists(tmp_path):
                            os.unlink(tmp_path)
                else:
                    zout.writestr(item, zin.read(item.filename))


def export_pptx_translation(original_path, translated, output_path):
    """全译文模式：保留格式，按页对位替换文字。

    优先使用 python-pptx（准确识别文本框/表格/组合形状结构）。
    仅在文件 >200MB 或 python-pptx 失败时回退到 zipfile 方式。
    """
    # 兜底：确保译文中有幻灯片标记
    translated = _reconstruct_slide_markers(original_path, translated)

    import gc
    import os

    file_size_mb = os.path.getsize(original_path) / (1024 * 1024)

    # 优先使用 python-pptx（准确），超大文件才用 zipfile
    if file_size_mb <= 200:
        try:
            _export_pptx_python_pptx(original_path, translated, output_path)
            return
        except Exception:
            pass

    # 回退：zipfile 方式（不加载图片，适合超大文件）
    try:
        gc.collect()
        _export_pptx_zip(original_path, translated, output_path)
        return
    except Exception:
        pass

    # 最后兜底：再试 python-pptx
    gc.collect()
    _export_pptx_python_pptx(original_path, translated, output_path)


def _export_pptx_python_pptx(original_path, translated, output_path):
    """使用 python-pptx 导出 PPT 翻译，准确识别文本结构。

    - 正确处理标题、正文、表格、组合形状
    - 表格按行匹配（与 read_pptx 的行级提取一致）
    - 使用 match_translation 智能对齐译文
    """
    import gc
    from pptx import Presentation

    gc.collect()
    trans_slides = _parse_translated_slides(translated)
    prs = Presentation(original_path)

    for slide_idx, slide in enumerate(prs.slides, 1):
        # 收集原始段落
        orig_paras = _collect_slide_paragraphs(slide)
        if not orig_paras:
            continue

        # 获取对应页的译文行
        trans_lines = trans_slides.get(slide_idx, [])
        if not trans_lines:
            continue

        # 分离标题和正文
        trans_title, trans_body_lines = _split_title_and_body(trans_lines)

        # ── 标题替换 ──
        title_paras = [p for p in orig_paras if p["is_title"]]
        if trans_title and title_paras:
            _replace_para_text(title_paras[0]["para"], trans_title)

        # ── 正文替换（含表格行） ──
        body_paras = [p for p in orig_paras if not p["is_title"]]
        if body_paras and trans_body_lines:
            orig_bodies = [p["text"] for p in body_paras]
            matched = match_translation(orig_bodies, "\n".join(trans_body_lines))

            for i, bp in enumerate(body_paras):
                if i >= len(matched):
                    break
                matched_text = matched[i]
                if not matched_text or not matched_text.strip():
                    # 译文为空，保留原文
                    continue

                if bp.get("is_table_row") and bp.get("row_cells"):
                    # 表格行：按 " | " 拆分译文，写回各单元格
                    cells = matched_text.split(" | ")
                    row_paras = bp["row_cells"]
                    for ci, cell_para in enumerate(row_paras):
                        if cell_para and ci < len(cells):
                            _replace_para_text(cell_para, cells[ci].strip())
                else:
                    # 普通段落
                    _replace_para_text(bp["para"], matched_text)

    prs.save(output_path)


def export_pptx_bilingual(original_path, translated, output_path):
    """双语备注模式：正文保留原文，译文写入演讲者备注栏。"""
    import gc
    from pptx import Presentation

    gc.collect()

    # 兜底：确保译文中有幻灯片标记
    translated = _reconstruct_slide_markers(original_path, translated)
    trans_slides = _parse_translated_slides(translated)
    prs = Presentation(original_path)

    for slide_idx, slide in enumerate(prs.slides, 1):
        trans_lines = trans_slides.get(slide_idx, [])
        if not trans_lines:
            continue

        # 写入备注
        try:
            notes = slide.notes_slide
        except Exception:
            slide.notes_slide  # 创建备注页
            notes = slide.notes_slide

        trans_title, trans_body = _split_title_and_body(trans_lines)
        notes_text = "【译文】\n"
        if trans_title:
            notes_text += f"标题：{trans_title}\n"
        notes_text += "\n".join(trans_body) if trans_body else "（无正文内容）"
        notes.notes_text_frame.text = notes_text

    prs.save(output_path)


def export_pptx_bilingual_inline(original_path, translated, output_path):
    """行内双语模式：在每个文本框中，原文下方追加译文行。

    视觉效果：
    ┌─────────────────────┐
    │ Original Text       │
    │ 原文翻译            │
    └─────────────────────┘

    使用 run 级别操作，保留原文格式，译文使用较小字号 + 灰色。
    """
    import gc
    from pptx import Presentation
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    gc.collect()

    # 兜底：确保译文中有幻灯片标记
    translated = _reconstruct_slide_markers(original_path, translated)
    trans_slides = _parse_translated_slides(translated)
    prs = Presentation(original_path)

    for slide_idx, slide in enumerate(prs.slides, 1):
        trans_lines = trans_slides.get(slide_idx, [])
        if not trans_lines:
            continue

        trans_title, trans_body_lines = _split_title_and_body(trans_lines)

        # 收集原始段落
        orig_paras = _collect_slide_paragraphs(slide)
        if not orig_paras:
            continue

        body_paras = [p for p in orig_paras if not p["is_title"]]
        orig_bodies = [p["text"] for p in body_paras]

        # 按比例对位
        if body_paras and trans_body_lines:
            matched = match_translation(orig_bodies, "\n".join(trans_body_lines))

            for i, bp in enumerate(body_paras):
                if i < len(matched) and matched[i]:
                    _append_translation_run(bp["para"], matched[i])

        # 标题：在标题段落后追加译文
        title_paras = [p for p in orig_paras if p["is_title"]]
        if trans_title and title_paras:
            _append_translation_run(title_paras[0]["para"], trans_title)

    prs.save(output_path)


def _append_translation_run(para, translated_text):
    """在段落后追加一个译文 run，使用较小字号 + 灰色。"""
    try:
        from pptx.util import Pt
        from pptx.dml.color import RGBColor

        # 获取原文 run 的字号作为参考
        ref_size = 10
        if para.runs:
            try:
                ref_size = para.runs[0].font.size.pt if para.runs[0].font.size else 10
            except Exception:
                pass

        # 添加换行 + 译文 run
        run = para.add_run()
        run.text = "\n" + translated_text
        try:
            run.font.size = Pt(max(ref_size - 2, 6))
            run.font.color.rgb = RGBColor(0x4E, 0x59, 0x69)  # Arco TEXT_REGULAR
        except Exception:
            pass
    except Exception:
        pass


# ─── 导出 Excel ──────────────────────────────────────────

def export_excel_bilingual(original_rows, translated, output_path, original_path=None):
    import openpyxl
    import copy
    from openpyxl.styles import PatternFill, Font

    if original_path and os.path.exists(original_path):
        wb = openpyxl.load_workbook(original_path)
    else:
        wb = openpyxl.Workbook()

    ws = wb.active
    trans_rows = [r for r in translated.split("\n") if r.strip()]
    max_col = ws.max_column

    header = ws.cell(row=1, column=max_col + 1, value="译文")
    header.fill = PatternFill("solid", fgColor="4F8EF7")
    header.font = Font(color="FFFFFF", bold=True)

    for i in range(2, ws.max_row + 1):
        trans_text = trans_rows[i - 2] if (i - 2) < len(trans_rows) else ""
        src_cell = ws.cell(row=i, column=max_col)
        new_cell = ws.cell(row=i, column=max_col + 1, value=trans_text)
        try:
            new_cell.font = copy.copy(src_cell.font)
            new_cell.alignment = copy.copy(src_cell.alignment)
            new_cell.border = copy.copy(src_cell.border)
        except Exception:
            pass

    wb.save(output_path)


def export_excel_translation(translated, output_path, original_path=None):
    import openpyxl
    import copy

    if original_path and os.path.exists(original_path):
        wb = openpyxl.load_workbook(original_path)
        ws = wb.active
        trans_rows = [r for r in translated.split("\n") if r.strip()]
        row_idx = 0
        for row in ws.iter_rows():
            if any(cell.value is not None for cell in row):
                row_text = trans_rows[row_idx] if row_idx < len(trans_rows) else ""
                cells_text = row_text.split("  ")
                for c_idx, cell in enumerate(row):
                    if cell.value is not None:
                        cell.value = cells_text[c_idx] if c_idx < len(cells_text) else ""
                row_idx += 1
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        for i, line in enumerate(translated.split("\n"), start=1):
            if line.strip():
                for j, val in enumerate(line.split("  "), start=1):
                    ws.cell(row=i, column=j, value=val)

    wb.save(output_path)


# ─── 导出 PDF ────────────────────────────────────────────

def register_fonts():
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        # 按优先级搜索 CJK 字体路径
        font_paths = [
            # Windows
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simsun.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            # Linux
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
            "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
            # macOS
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
        ]
        for path in font_paths:
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont("CJK", path))
                return "CJK"
    except Exception:
        pass
    return "Helvetica"


def export_pdf_translation(translated, output_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    font = register_fonts()
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    style = ParagraphStyle("b", fontName=font, fontSize=11,
                           leading=18, wordWrap="CJK")
    story = []
    for para in translated.split("\n"):
        if para.strip():
            story.append(Paragraph(para, style))
            story.append(Spacer(1, 6))
    doc.build(story)


def export_pdf_paragraph(original, translated, output_path):
    """段落对照PDF：一段原文一段译文"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib import colors
    font = register_fonts()
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    orig_style = ParagraphStyle(
        "orig", fontName=font, fontSize=10, leading=16,
        wordWrap="CJK", textColor=colors.HexColor("#555555"),
        backColor=colors.HexColor("#F5F5F5"),
        borderPadding=(6, 8, 6, 8),
    )
    trans_style = ParagraphStyle(
        "trans", fontName=font, fontSize=10, leading=16,
        wordWrap="CJK", textColor=colors.HexColor("#1A1A1A"),
        leftIndent=10,
        borderLeftWidth=3,
        borderLeftColor=colors.HexColor("#4F8EF7"),
        borderLeftPadding=8,
    )
    orig_paras = [p for p in original.split("\n") if p.strip()]
    matched_trans = match_translation(orig_paras, translated)
    story = []
    for i, orig_text in enumerate(orig_paras):
        story.append(Paragraph(orig_text, orig_style))
        story.append(Spacer(1, 4))
        trans_text = matched_trans[i] if i < len(matched_trans) else ""
        story.append(Paragraph(trans_text, trans_style))
        story.append(Spacer(1, 4))
        story.append(HRFlowable(width="100%", thickness=0.5,
                                color=colors.HexColor("#DDDDDD")))
        story.append(Spacer(1, 8))
    doc.build(story)


def export_pdf_bilingual(original, translated, output_path):
    """双列PDF"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
    from reportlab.lib import colors
    font = register_fonts()
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    style = ParagraphStyle("b", fontName=font, fontSize=9,
                           leading=14, wordWrap="CJK")
    hstyle = ParagraphStyle("h", fontName=font, fontSize=10, leading=14,
                            textColor=colors.HexColor("#4F8EF7"), wordWrap="CJK")
    orig_paras = [p for p in original.split("\n") if p.strip()]
    matched_trans = match_translation(orig_paras, translated)
    data = [[Paragraph("原文", hstyle), Paragraph("译文", hstyle)]]
    for i, orig_text in enumerate(orig_paras):
        t = matched_trans[i] if i < len(matched_trans) else ""
        data.append([Paragraph(orig_text, style), Paragraph(t, style)])
    col_w = (A4[0] - 3*cm) / 2
    tbl = Table(data, colWidths=[col_w, col_w], repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF4FF")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#F9F9F9")]),
    ]))
    doc.build([tbl])