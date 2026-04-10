"""
字幕处理模块
============
支持格式: SRT, VTT, ASS/SSA
功能: 解析字幕 → 翻译文本 → 输出双语字幕 / clean 纯文本
时间轴完整保留，格式正确，可直接用于视频。
"""

import re
import os
from datetime import timedelta
import re
import html as html_module


def clean_subtitle_entry(text: str) -> str:
    """清洗单条字幕文本"""
    # 去除 HTML 标签
    text = re.sub(r'<[^>]+>', '', text)
    # 去除 ASS/SSA 样式标签 {\xxx}
    text = re.sub(r'\{\\[^}]*\}', '', text)
    # 去除残留空大括号
    text = re.sub(r'\{\}', '', text)
    # 去除音乐符号 ♪
    text = text.replace('♪', '').replace('♫', '')
    # 去除全角空格
    text = text.replace('\u3000', ' ')
    # 合并多个空格为一个
    text = re.sub(r' {2,}', ' ', text)
    # 去除首尾空白
    text = text.strip()
    return text


def clean_subtitle_file(subtitle) -> object:
    """
    清洗字幕文件：
    1. 去除头部信息、注释、HTML标签、样式声明
    2. 合并相邻重复行
    3. 时间统一转为 SRT 格式 (00:00:00,000)
    4. 重编号

    参数: subtitle - SubtitleFile 对象
    返回: 清洗后的 SubtitleFile 对象（原地修改并返回）
    """
    if not subtitle or not subtitle.entries:
        return subtitle

    # ── 1. 清洗每条字幕文本 ──
    for entry in subtitle.entries:
        entry.text = clean_subtitle_entry(entry.text)

    # ── 2. 合并相邻重复行 ──
    merged = []
    prev_text = None
    for entry in subtitle.entries:
        if entry.text == prev_text and entry.text.strip() == '':
            # 跳过连续的空行
            continue
        if entry.text.strip() == '':
            merged.append(entry)
            prev_text = entry.text
            continue
        # 合并文本完全相同的相邻行（非空），合并时间范围
        if prev_text is not None and entry.text == prev_text and merged:
            # 扩展上一条的结束时间
            prev = merged[-1]
            prev.end_time = entry.end_time
        else:
            merged.append(entry)
            prev_text = entry.text
    subtitle.entries = merged

    # ── 3. 去除纯注释行（以 # 开头或 <i> 等残留） ──
    subtitle.entries = [
        e for e in subtitle.entries
        if e.text.strip() != '' or e.start_time != e.end_time
    ]

    # ── 4. 时间统一为 SRT 格式 (HH:MM:SS,mmm) ──
    for entry in subtitle.entries:
        entry.start_time = _normalize_timestamp(entry.start_time)
        entry.end_time = _normalize_timestamp(entry.end_time)

    # ── 5. 重编号 ──
    for i, entry in enumerate(subtitle.entries, start=1):
        entry.index = i

    return subtitle


def _normalize_timestamp(ts: str) -> str:
    """
    将各种时间戳格式统一为 SRT 格式: HH:MM:SS,mmm

    支持输入:
    - SRT:     00:01:23,456
    - VTT:     00:01:23.456
    - ASS:     0:01:23.45 或 H:MM:SS.cc
    """
    ts = ts.strip()

    # 匹配各种时间格式
    # 格式1: H:MM:SS.cc (ASS, 厘秒) 或 H:MM:SS.ccc
    # 格式2: HH:MM:SS,mmm (SRT, 毫秒)
    # 格式3: HH:MM:SS.mmm (VTT, 毫秒)

    # 统一替换：把逗号分隔符替换为点
    ts = ts.replace(',', '.')

    # 解析 ASS 格式: H:MM:SS.cc (2位厘秒)
    m = re.match(r'^(\d{1,2}):(\d{2}):(\d{2})\.(\d{1,3})$', ts)
    if m:
        h = int(m.group(1))
        mins = int(m.group(2))
        secs = int(m.group(3))
        frac = m.group(4)
        # 如果是2位厘秒，补零变3位毫秒
        if len(frac) == 2:
            frac = frac + '0'
        elif len(frac) == 1:
            frac = frac + '00'
        ms = int(frac[:3])
        return f"{h:02d}:{mins:02d}:{secs:02d},{ms:03d}"

    # 匹配 SRT/VTT 格式: HH:MM:SS.mmm
    m = re.match(r'^(\d{2}):(\d{2}):(\d{2})\.(\d{3})$', ts)
    if m:
        h = int(m.group(1))
        mins = int(m.group(2))
        secs = int(m.group(3))
        ms = int(m.group(4))
        return f"{h:02d}:{mins:02d}:{secs:02d},{ms:03d}"

    # 兜底：原样返回
    return ts.replace('.', ',')


# ═══════════════════════════════════════════════════════════
# 数据结构
# ═══════════════════════════════════════════════════════════

class SubtitleEntry:
    """单条字幕：序号 + 时间轴 + 文本"""
    def __init__(self, index, start_time, end_time, text,
                 style="", effect="", layer=0, margin_l=0, margin_r=0, margin_v=0):
        self.index = index
        self.start_time = start_time  # 格式: "00:01:23,456" (SRT) 或 "0:01:23.45" (ASS)
        self.end_time = end_time
        self.text = text
        self.translated = ""        # 译文（翻译后填充）
        # ASS 专用字段
        self.style = style
        self.effect = effect
        self.layer = layer
        self.margin_l = margin_l
        self.margin_r = margin_r
        self.margin_v = margin_v


class SubtitleFile:
    """字幕文件容器"""
    def __init__(self, format_type="srt", header=""):
        self.format_type = format_type  # "srt" | "vtt" | "ass"
        self.entries = []
        self.header = header

    @property
    def text_only(self):
        """获取所有字幕文本（纯文本，用于翻译）"""
        return "\n".join(e.text for e in self.entries if e.text.strip())

    @property
    def text_blocks(self):
        """获取文本块列表（每条字幕一个块）"""
        return [e.text for e in self.entries if e.text.strip()]

    @property
    def entry_count(self):
        return len(self.entries)


# ═══════════════════════════════════════════════════════════
# 时间轴工具
# ═══════════════════════════════════════════════════════════

def _srt_time_to_ms(time_str):
    """SRT 时间 -> 毫秒  '00:01:23,456' -> 83456"""
    h, m, s_ms = time_str.strip().split(":")
    s, ms = s_ms.split(",")
    return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)


def _ms_to_srt_time(ms):
    """毫秒 -> SRT 时间"""
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms %= 1000
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def _vtt_time_to_ms(time_str):
    """VTT 时间 -> 毫秒  '00:01:23.456' -> 83456"""
    time_str = time_str.strip()
    parts = time_str.replace(".", ":").split(":")
    if len(parts) == 3:
        h, m, s = parts
    else:
        h, m, s = "0", parts[0], parts[1]
    return int(h) * 3600000 + int(m) * 60000 + int(float(s) * 1000)


def _ms_to_vtt_time(ms):
    """毫秒 -> VTT 时间  MM:SS.mmm"""
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms %= 1000
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}.{ms:03d}"
    return f"{m:02d}:{s:02d}.{ms:03d}"


def _ass_time_to_ms(time_str):
    """ASS 时间 -> 毫秒  '0:01:23.45' -> 83450"""
    h, m, s_cs = time_str.strip().split(":")
    s, cs = s_cs.split(".")
    return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(cs) * 10


def _ms_to_ass_time(ms):
    """毫秒 -> ASS 时间  H:MM:SS.cc"""
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    cs = (ms % 1000) // 10
    return f"{h}:{m:02d}:{s:02d}.{cs:02d}"


# ═══════════════════════════════════════════════════════════
# 解析器
# ═══════════════════════════════════════════════════════════

def parse_srt(text):
    """解析 SRT 字幕文件"""
    sub = SubtitleFile(format_type="srt")
    blocks = re.split(r"\n\s*\n", text.strip())
    index = 0

    for block in blocks:
        lines = block.strip().split("\n")
        if len(lines) < 2:
            continue

        time_line_idx = -1
        for i, line in enumerate(lines):
            if "-->" in line:
                time_line_idx = i
                break

        if time_line_idx < 0:
            continue

        time_line = lines[time_line_idx]
        match = re.match(
            r"(\d{2}:\d{2}:\d{2}[,\.]\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}[,\.]\d{3})",
            time_line.strip()
        )
        if not match:
            continue

        start = match.group(1).replace(".", ",")
        end = match.group(2).replace(".", ",")

        if time_line_idx > 0 and lines[time_line_idx - 1].strip().isdigit():
            seq = int(lines[time_line_idx - 1].strip())
        else:
            seq = index + 1

        text_lines = lines[time_line_idx + 1:]
        text = "\n".join(text_lines).strip()
        text = re.sub(r"</?[^>]+>", "", text)

        if text.strip():
            index += 1
            sub.entries.append(SubtitleEntry(
                index=seq, start_time=start, end_time=end, text=text
            ))

    return sub


def parse_vtt(text):
    """解析 WebVTT 字幕文件"""
    sub = SubtitleFile(format_type="vtt", header="WEBVTT")

    lines = text.strip().split("\n")
    header_lines = []
    content_start = 0
    for i, line in enumerate(lines):
        if line.strip().startswith("WEBVTT"):
            header_lines.append(line.strip())
        elif line.strip().startswith("NOTE") or line.strip().startswith("Kind:") or \
             line.strip().startswith("Language:"):
            header_lines.append(line.strip())
        elif header_lines and not line.strip():
            content_start = i + 1
            break
        elif not header_lines and "-->" in line:
            content_start = i
            break

    sub.header = "\n".join(header_lines) if header_lines else "WEBVTT"

    content = "\n".join(lines[content_start:])
    blocks = re.split(r"\n\s*\n", content.strip())
    index = 0

    for block in blocks:
        block_lines = block.strip().split("\n")
        if not block_lines:
            continue

        time_line_idx = -1
        for i, line in enumerate(block_lines):
            if "-->" in line:
                time_line_idx = i
                break

        if time_line_idx < 0:
            continue

        time_line = block_lines[time_line_idx]
        match = re.match(
            r"(\d{1,2}:?\d{2}:\d{2}\.\d{3})\s*-->\s*(\d{1,2}:?\d{2}:\d{2}\.\d{3})",
            time_line.strip()
        )
        if not match:
            continue

        start = match.group(1)
        end = match.group(2)

        text_lines = block_lines[time_line_idx + 1:]
        text = "\n".join(text_lines).strip()
        text = re.sub(r"</?[^>]+>", "", text)

        if text.strip():
            index += 1
            sub.entries.append(SubtitleEntry(
                index=index, start_time=start, end_time=end, text=text
            ))

    return sub


def parse_ass(text):
    """解析 ASS/SSA 字幕文件"""
    sub = SubtitleFile(format_type="ass")

    lines = text.split("\n")
    format_fields = None
    header_lines = []

    for line in lines:
        stripped = line.strip()

        if stripped.startswith("[Events]"):
            header_lines.append(line)
            continue
        elif not format_fields and (stripped.startswith("[") or stripped.startswith(";") or
                                     stripped.startswith("ScriptType") or stripped.startswith("PlayRes") or
                                     stripped.startswith("WrapStyle") or stripped.startswith("ScaledBorderAndShadow")):
            header_lines.append(line)
            continue

        if stripped.startswith("Format:"):
            format_fields = [f.strip() for f in stripped[len("Format:"):].split(",")]
            header_lines.append(line)
            continue

        if stripped.startswith("Dialogue:"):
            if not format_fields:
                format_fields = ["Layer", "Start", "End", "Style", "Name",
                                 "MarginL", "MarginR", "MarginV", "Effect", "Text"]

            content = stripped[len("Dialogue:"):].strip()
            parts = []
            current = ""
            in_brace = 0
            for ch in content:
                if ch == '{':
                    in_brace += 1
                    current += ch
                elif ch == '}':
                    in_brace = max(0, in_brace - 1)
                    current += ch
                elif ch == ',' and in_brace == 0 and len(parts) < 9:
                    parts.append(current.strip())
                    current = ""
                else:
                    current += ch
            parts.append(current.strip())

            field_map = {}
            for i, field_name in enumerate(format_fields):
                if i < len(parts):
                    field_map[field_name] = parts[i]

            start = field_map.get("Start", "0:00:00.00")
            end = field_map.get("End", "0:00:00.00")
            dialog_text = field_map.get("Text", "")

            clean_text = re.sub(r"\{[^}]*\}", "", dialog_text)
            clean_text = clean_text.replace("\\N", "\n").replace("\\n", "\n")

            if clean_text.strip():
                sub.entries.append(SubtitleEntry(
                    index=len(sub.entries) + 1,
                    start_time=start,
                    end_time=end,
                    text=clean_text,
                    style=field_map.get("Style", "Default"),
                    effect=field_map.get("Effect", ""),
                    layer=int(field_map.get("Layer", "0")),
                    margin_l=int(field_map.get("MarginL", "0")),
                    margin_r=int(field_map.get("MarginR", "0")),
                    margin_v=int(field_map.get("MarginV", "0")),
                ))
        else:
            if not format_fields:
                header_lines.append(line)

    sub.header = "\n".join(header_lines)
    return sub


# ═══════════════════════════════════════════════════════════
# 自动识别格式 + 读取
# ═══════════════════════════════════════════════════════════

def detect_format(text):
    """根据内容自动检测字幕格式"""
    first_line = text.strip().split("\n")[0].strip()
    if first_line.startswith("WEBVTT"):
        return "vtt"
    if first_line.startswith("[Script Info]") or "Dialogue:" in text[:2000]:
        return "ass"
    if re.search(r"\d+\s*\n\s*\d{2}:\d{2}:\d{2}[,\.]\d{3}\s*-->", text[:1000]):
        return "srt"
    return "srt"


def parse_subtitle(text, format_type=None):
    """解析字幕文本，自动识别或指定格式"""
    if not format_type:
        format_type = detect_format(text)
    format_type = format_type.lower().strip()

    if format_type == "vtt":
        return parse_vtt(text)
    elif format_type in ("ass", "ssa"):
        return parse_ass(text)
    else:
        return parse_srt(text)


def read_subtitle_file(file_path):
    """读取字幕文件"""
    with open(file_path, "r", encoding="utf-8-sig") as f:
        text = f.read()

    ext = os.path.splitext(file_path)[1].lower().lstrip(".")
    format_map = {"srt": "srt", "vtt": "vtt", "ass": "ass", "ssa": "ass"}
    fmt = format_map.get(ext, None)

    return parse_subtitle(text, fmt)


# ═══════════════════════════════════════════════════════════
# 译文填充
# ═══════════════════════════════════════════════════════════

def _apply_translations(subtitle, translated_texts):
    """将翻译结果对位填充到字幕条目中"""
    trans_idx = 0
    for entry in subtitle.entries:
        if not entry.text.strip():
            continue
        if trans_idx < len(translated_texts):
            entry.translated = translated_texts[trans_idx].strip()
            trans_idx += 1


# ═══════════════════════════════════════════════════════════
# 时间轴格式转换
# ═══════════════════════════════════════════════════════════

def _normalize_time_to_srt(time_str, fmt="srt"):
    try:
        if fmt == "ass":
            ms = _ass_time_to_ms(time_str)
        elif fmt == "vtt":
            ms = _vtt_time_to_ms(time_str)
        else:
            ms = _srt_time_to_ms(time_str)
        return _ms_to_srt_time(ms)
    except Exception:
        return time_str


def _normalize_time_to_vtt(time_str, fmt="srt"):
    try:
        if fmt == "ass":
            ms = _ass_time_to_ms(time_str)
        elif fmt == "vtt":
            ms = _vtt_time_to_ms(time_str)
        else:
            ms = _srt_time_to_ms(time_str)
        return _ms_to_vtt_time(ms)
    except Exception:
        return time_str


def _normalize_time_to_ass(time_str, fmt="srt"):
    try:
        if fmt == "ass":
            ms = _ass_time_to_ms(time_str)
        elif fmt == "vtt":
            ms = _vtt_time_to_ms(time_str)
        else:
            ms = _srt_time_to_ms(time_str)
        return _ms_to_ass_time(ms)
    except Exception:
        return time_str


# ═══════════════════════════════════════════════════════════
# 导出器
# ═══════════════════════════════════════════════════════════

def export_srt_bilingual(subtitle, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        for i, entry in enumerate(subtitle.entries):
            start = _normalize_time_to_srt(entry.start_time, subtitle.format_type)
            end = _normalize_time_to_srt(entry.end_time, subtitle.format_type)
            f.write(f"{i + 1}\n")
            f.write(f"{start} --> {end}\n")
            f.write(f"{entry.text}\n")
            if entry.translated:
                f.write(f"{entry.translated}\n")
            f.write("\n")


def export_srt_translated(subtitle, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        for i, entry in enumerate(subtitle.entries):
            text = entry.translated if entry.translated else entry.text
            start = _normalize_time_to_srt(entry.start_time, subtitle.format_type)
            end = _normalize_time_to_srt(entry.end_time, subtitle.format_type)
            f.write(f"{i + 1}\n")
            f.write(f"{start} --> {end}\n")
            f.write(f"{text}\n")
            f.write("\n")


def export_vtt_bilingual(subtitle, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("WEBVTT\n\n")
        for entry in subtitle.entries:
            start = _normalize_time_to_vtt(entry.start_time, subtitle.format_type)
            end = _normalize_time_to_vtt(entry.end_time, subtitle.format_type)
            f.write(f"{start} --> {end}\n")
            f.write(f"{entry.text}\n")
            if entry.translated:
                f.write(f"{entry.translated}\n")
            f.write("\n")


def export_vtt_translated(subtitle, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("WEBVTT\n\n")
        for entry in subtitle.entries:
            text = entry.translated if entry.translated else entry.text
            start = _normalize_time_to_vtt(entry.start_time, subtitle.format_type)
            end = _normalize_time_to_vtt(entry.end_time, subtitle.format_type)
            f.write(f"{start} --> {end}\n")
            f.write(f"{text}\n")
            f.write("\n")


def export_ass_bilingual(subtitle, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        if subtitle.header:
            f.write(subtitle.header + "\n")

        f.write("\n[Events]\n")
        f.write("Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n")

        for entry in subtitle.entries:
            text = entry.text
            if entry.translated:
                text = entry.text + "\\N" + entry.translated

            text = text.replace("\n", "\\N")

            start = _normalize_time_to_ass(entry.start_time, subtitle.format_type)
            end = _normalize_time_to_ass(entry.end_time, subtitle.format_type)

            f.write(
                f"Dialogue: {entry.layer},{start},{end},"
                f"{entry.style},,{entry.margin_l},{entry.margin_r},{entry.margin_v},,"
                f"{text}\n"
            )


def export_ass_translated(subtitle, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        if subtitle.header:
            f.write(subtitle.header + "\n")

        f.write("\n[Events]\n")
        f.write("Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n")

        for entry in subtitle.entries:
            text = entry.translated if entry.translated else entry.text
            text = text.replace("\n", "\\N")

            start = _normalize_time_to_ass(entry.start_time, subtitle.format_type)
            end = _normalize_time_to_ass(entry.end_time, subtitle.format_type)

            f.write(
                f"Dialogue: {entry.layer},{start},{end},"
                f"{entry.style},,{entry.margin_l},{entry.margin_r},{entry.margin_v},,"
                f"{text}\n"
            )


def export_clean_text(subtitle, output_path, bilingual=False):
    with open(output_path, "w", encoding="utf-8") as f:
        for entry in subtitle.entries:
            if bilingual:
                f.write(f"{entry.text}\n")
                if entry.translated:
                    f.write(f"{entry.translated}\n")
                f.write("\n")
            else:
                text = entry.translated if entry.translated else entry.text
                f.write(f"{text}\n")


def export_subtitle(subtitle, output_path, mode="bilingual"):
    """通用导出函数

    Args:
        subtitle: SubtitleFile
        output_path: 输出文件路径
        mode: "bilingual" | "translated" | "clean" | "clean_bilingual"
    """
    ext = os.path.splitext(output_path)[1].lower()

    if ext in (".ass", ".ssa"):
        target_fmt = "ass"
    elif ext == ".vtt":
        target_fmt = "vtt"
    else:
        target_fmt = "srt"

    if mode == "bilingual":
        if target_fmt == "ass":
            export_ass_bilingual(subtitle, output_path)
        elif target_fmt == "vtt":
            export_vtt_bilingual(subtitle, output_path)
        else:
            export_srt_bilingual(subtitle, output_path)
    elif mode == "translated":
        if target_fmt == "ass":
            export_ass_translated(subtitle, output_path)
        elif target_fmt == "vtt":
            export_vtt_translated(subtitle, output_path)
        else:
            export_srt_translated(subtitle, output_path)
    elif mode == "clean":
        export_clean_text(subtitle, output_path, bilingual=False)
    elif mode == "clean_bilingual":
        export_clean_text(subtitle, output_path, bilingual=True)
    else:
        export_srt_bilingual(subtitle, output_path)

    return output_path


# ═══════════════════════════════════════════════════════════
# 字幕翻译提示词
# ═══════════════════════════════════════════════════════════

def build_subtitle_translate_prompt(source_lang, target_lang):
    """构建字幕翻译专用提示词"""
    return f"""You are a professional subtitle translator translating from {source_lang} to {target_lang}.

CRITICAL RULES:
1. Each input line is ONE subtitle entry. Translate EACH line independently.
2. Output MUST have the EXACT SAME number of lines as input. No more, no fewer.
3. Keep translations concise and natural - suitable for on-screen subtitle reading.
4. Do NOT add explanations, notes, or comments.
5. Do NOT merge or split lines.
6. For empty lines, output an empty line.
7. Preserve the tone and style of the original (formal, casual, humorous, etc.).
8. If a line contains only a sound effect or non-verbal text (e.g., "(laughs)", "[music]"), translate it naturally.
"""


def format_subtitle_for_translation(subtitle):
    """将字幕格式化为翻译输入文本"""
    lines = []
    for i, entry in enumerate(subtitle.entries):
        if entry.text.strip():
            lines.append(f"[{i+1:03d}] {entry.text}")
    return "\n".join(lines)


def parse_subtitle_translation_response(response_text, expected_count):
    """解析 AI 翻译结果"""
    translations = []

    # 方法1: 匹配 [序号] 译文
    pattern = re.compile(r"\[(\d{3})\]\s*(.+)")
    matches = pattern.findall(response_text)
    if matches and len(matches) >= expected_count * 0.5:
        ordered = sorted(matches, key=lambda x: int(x[0]))
        translations = [m[1].strip() for m in ordered]
        while len(translations) < expected_count:
            translations.append("")
        return translations[:expected_count]

    # 方法2: 直接按行分割
    lines = [l.strip() for l in response_text.strip().split("\n") if l.strip()]
    if len(lines) >= expected_count * 0.5:
        return lines[:expected_count]

    # 方法3: 去除序号前缀后匹配
    cleaned_lines = []
    for line in response_text.strip().split("\n"):
        line = re.sub(r"^\s*\[\d+\]\s*", "", line).strip()
        if line:
            cleaned_lines.append(line)
    if len(cleaned_lines) >= expected_count * 0.5:
        return cleaned_lines[:expected_count]

    return translations if translations else lines


# ═══════════════════════════════════════════════════════════
# 文件名生成
# ═══════════════════════════════════════════════════════════

def generate_subtitle_output_name(original_path, mode="bilingual",
                                  source_lang="English", target_lang="Chinese"):
    """生成字幕输出文件名

    格式: 原文件名-EC/CE-agent.扩展名
    EC = English->Chinese, CE = Chinese->English
    """
    basename = os.path.splitext(os.path.basename(original_path))[0]
    ext = os.path.splitext(original_path)[1]

    lang_code_map = {"english": "E", "chinese": "C", "japanese": "J", "korean": "K",
                     "french": "F", "german": "G", "spanish": "S"}
    sl = lang_code_map.get(source_lang.lower(), source_lang[:1].upper())
    tl = lang_code_map.get(target_lang.lower(), target_lang[:1].upper())
    lang_suffix = f"{sl}{tl}"

    mode_suffix = {"bilingual": "", "translated": "-trans",
                   "clean": "-clean", "clean_bilingual": "-clean"}
    mode_str = mode_suffix.get(mode, "")

    if mode in ("clean", "clean_bilingual"):
        ext = ".txt"

    return f"{basename}-{lang_suffix}-agent{mode_str}{ext}"
# ═══════════════════════════════════════════════════════════
# 兼容别名（供 gui.py 调用）
# ═══════════════════════════════════════════════════════════
parse_subtitle_file = read_subtitle_file


def export_subtitle_file(subtitle, output_path, output_mode="bilingual", output_format=None):
    """兼容接口：导出字幕文件，output_format 由文件扩展名自动判断"""
    return export_subtitle(subtitle, output_path, mode=output_mode)
# ═══════════════════════════════════════════════════════════
# 字幕视频适配优化
# ═══════════════════════════════════════════════════════════

def _count_cjk(text):
    """统计中日韩字符数"""
    return sum(1 for c in text if '\u4e00' <= c <= '\u9fff' or '\u3040' <= c <= '\u309f' or '\u30a0' <= c <= '\u30ff')


def _is_cjk_dominant(text):
    """判断文本是否以中日韩文字为主"""
    cjk = _count_cjk(text)
    return cjk > len(text) * 0.3


def _smart_line_break(text, max_chars=42, is_cjk=False):
    """
    智能换行：在自然停顿处断行
    返回最多2行，每行不超过 max_chars
    """
    if len(text) <= max_chars:
        return text

    if is_cjk:
        max_chars = min(max_chars, 20)
    else:
        max_chars = min(max_chars, 42)

    if len(text) <= max_chars:
        return text

    # 优先在标点处断行
    break_chars_cjk = "，。！？；：、""''）】》"
    break_chars_en = ",.!?;:)'\"}]>"
    break_chars = break_chars_cjk + break_chars_en

    best_pos = -1
    for i, c in enumerate(text):
        if c in break_chars and i < len(text) - 1:
            # CJK: 断在标点后面
            if c in break_chars_cjk:
                best_pos = i + 1
            # EN: 断在标点后面（空格处）
            elif c in break_chars_en:
                best_pos = i + 1

        if best_pos > 0 and best_pos >= max_chars * 0.6:
            break

    if best_pos < 0:
        # 没有标点，按字数对半切
        mid = len(text) // 2
        # 找最近的空格
        space_pos = text.rfind(' ', 0, mid + 5)
        if space_pos > len(text) * 0.3:
            best_pos = space_pos + 1
        else:
            best_pos = mid

    line1 = text[:best_pos].strip()
    line2 = text[best_pos:].strip()

    # 如果第二行仍然太长，强制截断
    if len(line2) > max_chars and not is_cjk:
        line2 = line2[:max_chars - 3] + "..."

    return line1 + "\n" + line2


def _ms_from_srt(time_str):
    """SRT 时间字符串 -> 毫秒"""
    try:
        time_str = time_str.strip()
        if ',' in time_str:
            h, m, s_ms = time_str.split(":")
            s, ms = s_ms.split(",")
            return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)
        else:
            h, m, s_ms = time_str.split(":")
            s, ms = s_ms.split(".")
            return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)
    except Exception:
        return 0


def _ms_to_srt(ms):
    """毫秒 -> SRT 时间字符串"""
    ms = max(0, int(ms))
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms_r = ms % 1000
    return f"{h:02d}:{m:02d}:{s:02d},{ms_r:03d}"


def _ms_to_vtt(ms):
    """毫秒 -> VTT 时间字符串"""
    ms = max(0, int(ms))
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms_r = ms % 1000
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}.{ms_r:03d}"
    return f"{m:02d}:{s:02d}.{ms_r:03d}"


def _ms_to_ass(ms):
    """毫秒 -> ASS 时间字符串"""
    ms = max(0, int(ms))
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    cs = (ms % 1000) // 10
    return f"{h}:{m:02d}:{s:02d}.{cs:02d}"


def optimize_subtitle_for_video(subtitle, target_cps=15, max_lines=2,
                                 min_duration_ms=1000, max_duration_ms=7000,
                                 min_gap_ms=200):
    """
    优化字幕以适配视频播放

    Args:
        subtitle: SubtitleFile 对象
        target_cps: 目标每秒字符数（中文15，英文25）
        max_lines: 每条字幕最大行数（默认2行）
        min_duration_ms: 最短显示时长（毫秒，默认1000）
        max_duration_ms: 最长显示时长（毫秒，默认7000）
        min_gap_ms: 条目之间最短间隔（毫秒，默认200）

    Returns:
        优化后的 SubtitleFile 对象
    """
    if not subtitle or not subtitle.entries:
        return subtitle

    entries = subtitle.entries
    total = len(entries)

    for i, entry in enumerate(entries):
        # 使用译文（如有），否则用原文
        text = entry.translated if entry.translated else entry.text
        if not text.strip():
            continue

        is_cjk = _is_cjk_dominant(text)
        cps = target_cps if is_cjk else 25

        # 1. 智能换行
        optimized_text = _smart_line_break(text, max_chars=20 if is_cjk else 42, is_cjk=is_cjk)
        if entry.translated:
            entry.translated = optimized_text
        else:
            entry.text = optimized_text

        # 重新获取换行后的文本长度用于时长计算
        display_text = optimized_text
        char_count = len(display_text.replace("\n", ""))

        # 2. 根据文本长度计算最佳时长
        # 时长 = 字符数 / CPS，但受限于最短和最长
        optimal_ms = int(char_count / cps * 1000)
        optimal_ms = max(min_duration_ms, min(optimal_ms, max_duration_ms))

        # 3. 解析当前时间
        start_ms = _ms_from_srt(entry.start_time)
        end_ms = _ms_from_srt(entry.end_time)
        current_duration = end_ms - start_ms

        if current_duration < min_duration_ms:
            # 时长不足，延长到最优时长
            end_ms = start_ms + optimal_ms
        elif current_duration > max_duration_ms:
            # 时长过长，截断到最长
            end_ms = start_ms + max_duration_ms

        # 4. 确保与下一条不重叠，留出最小间隔
        if i < total - 1:
            next_start_ms = _ms_from_srt(entries[i + 1].start_time)
            if end_ms + min_gap_ms > next_start_ms and next_start_ms > start_ms:
                end_ms = next_start_ms - min_gap_ms
                # 如果调整后时长太短，保持原样
                if end_ms - start_ms < min_duration_ms:
                    end_ms = start_ms + min_duration_ms

        # 5. 回写时间
        entry.start_time = _ms_to_srt(start_ms)
        entry.end_time = _ms_to_srt(end_ms)

    # 6. 重编号
    for i, entry in enumerate(entries, start=1):
        entry.index = i

    # ═══════════════════════════════════════════════════════════
    # 字幕样式自动调整
    # ═══════════════════════════════════════════════════════════

    # 预定义样式表：按文本长度自动匹配
    SUBTITLE_STYLES = {
        "short": {  # ≤ 10 字符：大号，适合标题/感叹词
            "fontname": "Microsoft YaHei",
            "fontsize": 48,
            "primarycolor": "&H00FFFFFF",  # 白色
            "secondarycolor": "&H000000FF",  # 黑色
            "outlinecolor": "&H00000000",  # 黑色描边
            "backcolor": "&H40000000",  # 半透明黑背景
            "bold": True,
            "italic": False,
            "underline": False,
            "strikeout": False,
            "scalex": 100,
            "scaley": 100,
            "spacing": 0,
            "angle": 0,
            "borderstyle": 1,  # 1=描边+阴影
            "outline": 3,
            "shadow": 2,
            "alignment": 8,  # 底部居中
            "marginl": 10,
            "marginr": 10,
            "marginv": 30,
        },
        "medium": {  # 11~25 字符：中号，适合普通对话
            "fontname": "Microsoft YaHei",
            "fontsize": 36,
            "primarycolor": "&H00FFFFFF",
            "secondarycolor": "&H000000FF",
            "outlinecolor": "&H00000000",
            "backcolor": "&H40000000",
            "bold": False,
            "italic": False,
            "underline": False,
            "strikeout": False,
            "scalex": 100,
            "scaley": 100,
            "spacing": 0,
            "angle": 0,
            "borderstyle": 1,
            "outline": 2,
            "shadow": 1,
            "alignment": 8,
            "marginl": 10,
            "marginr": 10,
            "marginv": 30,
        },
        "long": {  # 26~42 字符：小号，适合长句
            "fontname": "Microsoft YaHei",
            "fontsize": 28,
            "primarycolor": "&H00FFFFFF",
            "secondarycolor": "&H000000FF",
            "outlinecolor": "&H00000000",
            "backcolor": "&H40000000",
            "bold": False,
            "italic": False,
            "underline": False,
            "strikeout": False,
            "scalex": 100,
            "scaley": 100,
            "spacing": 0,
            "angle": 0,
            "borderstyle": 1,
            "outline": 2,
            "shadow": 1,
            "alignment": 8,
            "marginl": 10,
            "marginr": 10,
            "marginv": 30,
        },
        "tiny": {  # > 42 字符：最小号
            "fontname": "Microsoft YaHei",
            "fontsize": 22,
            "primarycolor": "&H00FFFFFF",
            "secondarycolor": "&H000000FF",
            "outlinecolor": "&H00000000",
            "backcolor": "&H40000000",
            "bold": False,
            "italic": False,
            "underline": False,
            "strikeout": False,
            "scalex": 100,
            "scaley": 100,
            "spacing": 0,
            "angle": 0,
            "borderstyle": 1,
            "outline": 2,
            "shadow": 1,
            "alignment": 8,
            "marginl": 10,
            "marginr": 10,
            "marginv": 30,
        },
    }

    def _detect_style_name(text, translated=None):
        """根据文本长度选择样式级别"""
        display = translated if translated else text
        display = display.replace("\n", "")
        char_count = len(display)

        if char_count <= 10:
            return "short"
        elif char_count <= 25:
            return "medium"
        elif char_count <= 42:
            return "long"
        else:
            return "tiny"

    def _build_ass_styles_section(styles=None):
        """生成 ASS [V4+ Styles] 段落"""
        if not styles:
            styles = SUBTITLE_STYLES

        lines = ["[V4+ Styles]"]
        lines.append("Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, "
                     "OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, "
                     "ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, "
                     "Alignment, MarginL, MarginR, MarginV, Encoding")

        for name, s in styles.items():
            bold = -1 if s["bold"] else 0
            italic = -1 if s["italic"] else 0
            lines.append(
                f"Style: {name},{s['fontname']},{s['fontsize']},"
                f"{s['primarycolor']},{s['secondarycolor']},"
                f"{s['outlinecolor']},{s['backcolor']},"
                f"{bold},{italic},{s['underline']},{s['strikeout']},"
                f"{s['scalex']},{s['scaley']},{s['spacing']},{s['angle']},"
                f"{s['borderstyle']},{s['outline']},{s['shadow']},"
                f"{s['alignment']},{s['marginl']},{s['marginr']},{s['marginv']},1"
            )

        return "\n".join(lines)

    def apply_auto_style(subtitle):
        """
        为每条字幕自动分配样式名称（基于文本长度）
        修改 entry.style 字段
        """
        if not subtitle or not subtitle.entries:
            return subtitle

        for entry in subtitle.entries:
            style_name = _detect_style_name(entry.text, entry.translated)
            entry.style = style_name

        return subtitle

    def build_ass_with_styles(subtitle):
        """
        生成带样式的完整 ASS 文件内容
        包含 [Script Info] + [V4+ Styles] + [Events]
        """
        # 确保每条有条目有样式
        apply_auto_style(subtitle)

        lines = []

        # ── [Script Info] ──
        lines.append("[Script Info]")
        lines.append("ScriptType: v4.00+")
        lines.append("PlayResX: 1920")
        lines.append("PlayResY: 1080")
        lines.append("WrapStyle: 0")
        lines.append("ScaledBorderAndShadow: yes")
        lines.append("")

        # ── [V4+ Styles] ──
        lines.append(_build_ass_styles_section())
        lines.append("")

        # ── [Events] ──
        lines.append("[Events]")
        lines.append("Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text")

        for entry in subtitle.entries:
            text = entry.text
            if entry.translated:
                text = entry.text + "\\N" + entry.translated
            text = text.replace("\n", "\\N")

            start = _ms_to_ass(_ms_from_srt(entry.start_time))
            end = _ms_to_ass(_ms_from_srt(entry.end_time))

            lines.append(
                f"Dialogue: {entry.layer},{start},{end},"
                f"{entry.style},,{entry.margin_l},{entry.margin_r},{entry.margin_v},,"
                f"{text}"
            )

        return "\n".join(lines) + "\n"

    def build_srt_with_inline_style(subtitle):
        """
        生成带内嵌样式的 SRT 文件
        使用 ASS 兼容标签 {\fs48}{\b1} 等（大部分播放器支持）
        """
        lines = []

        for i, entry in enumerate(subtitle.entries):
            style_name = _detect_style_name(entry.text, entry.translated)
            fontsize = SUBTITLE_STYLES[style_name]["fontsize"]
            bold_tag = "\\b1" if SUBTITLE_STYLES[style_name]["bold"] else ""

            start = entry.start_time
            end = entry.end_time

            # 原文
            src_text = f"{{\\fs{fontsize}}}{bold_tag}{entry.text}" if entry.text.strip() else ""
            # 译文
            tgt_text = ""
            if entry.translated:
                tgt_text = f"{{\\fs{fontsize}}}{bold_tag}{entry.translated}"

            lines.append(f"{i + 1}")
            lines.append(f"{start} --> {end}")
            if src_text:
                lines.append(src_text)
            if tgt_text and tgt_text != src_text:
                lines.append(tgt_text)
            lines.append("")

        return "\n".join(lines)

    def build_vtt_with_css_style(subtitle):
        """
        生成带 CSS 样式的 VTT 文件
        使用 STYLE 标签和 class 控制字体大小
        """
        lines = ["WEBVTT"]

        # ── CSS 样式 ──
        lines.append("")
        lines.append("STYLE")
        lines.append("::cue {")
        lines.append("  font-family: 'Microsoft YaHei', sans-serif;")
        lines.append("  color: white;")
        lines.append("  text-shadow: 1px 1px 2px black;")
        lines.append("  background-color: rgba(0,0,0,0.4);")
        lines.append("  padding: 4px 8px;")
        lines.append("  border-radius: 4px;")
        lines.append("}")
        lines.append("")
        lines.append("::cue(.short) { font-size: 2.2em; }")
        lines.append("::cue(.medium) { font-size: 1.6em; }")
        lines.append("::cue(.long) { font-size: 1.2em; }")
        lines.append("::cue(.tiny) { font-size: 0.95em; }")
        lines.append("")

        for entry in subtitle.entries:
            style_name = _detect_style_name(entry.text, entry.translated)
            start = _ms_to_vtt(_ms_from_srt(entry.start_time))
            end = _ms_to_vtt(_ms_from_srt(entry.end_time))

            # 用 voice tag 充当 class
            tag = f"<v {style_name}>"
            text = entry.text
            if entry.translated:
                text = entry.text + "\n" + entry.translated
            text = text.replace("\n", "\n" + tag)

            lines.append(f"{start} --> {end}")
            lines.append(f"{tag}{text}")
            lines.append("")

        return "\n".join(lines)

    # ── 覆盖原有的导出函数，使其带样式 ──

    def export_ass_styled(subtitle, output_path):
        """导出带自动样式的 ASS 字幕"""
        content = build_ass_with_styles(subtitle)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)

    def export_srt_styled(subtitle, output_path):
        """导出带内嵌样式的 SRT 字幕"""
        content = build_srt_with_inline_style(subtitle)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)

    def export_vtt_styled(subtitle, output_path):
        """导出带 CSS 样式的 VTT 字幕"""
        content = build_vtt_with_css_style(subtitle)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)
    return subtitle
