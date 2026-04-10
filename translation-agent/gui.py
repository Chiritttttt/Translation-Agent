import os
import sys
from dotenv import load_dotenv
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QTabWidget, QMessageBox, QGroupBox, QProgressBar, QDialog,
    QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
import openai
import requests
from bs4 import BeautifulSoup

load_dotenv()

client = openai.OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url=os.getenv("OPENAI_BASE_URL"),
)
model = os.getenv("OPENAI_MODEL", "deepseek-chat")


# ─── AI 调用 ─────────────────────────────────────────────
def chat(system, user, temperature=0.3):
    resp = client.chat.completions.create(
        model=model, temperature=temperature,
        messages=[{"role":"system","content":system},
                  {"role":"user","content":user}]
    )
    return resp.choices[0].message.content


def fetch_url(url):
    try:
        resp = requests.get(f"https://r.jina.ai/{url}", timeout=15)
        if resp.status_code == 200 and len(resp.text) > 200:
            return resp.text
    except Exception:
        pass
    try:
        resp = requests.get(url, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script","style","nav","footer"]):
            tag.decompose()
        text = soup.get_text(separator="\n").strip()
        if len(text) > 200:
            return text
    except Exception:
        pass
    raise RuntimeError("无法抓取URL，请手动复制正文。")


def split_chunks(text, max_words=3000):
    words = text.split()
    if len(words) <= max_words:
        return [text]
    chunks, current = [], []
    for word in words:
        current.append(word)
        if len(current) >= max_words:
            chunks.append(" ".join(current))
            current = []
    if current:
        chunks.append(" ".join(current))
    return chunks


# ─── 五步翻译工作流 ──────────────────────────────────────
def step1_analyze(text, source_lang, target_lang, audience, style):
    sys_prompt = """你是专业翻译前置分析师。
输出严格按照以下固定Markdown结构：
## 概要
## 术语表（原文 → 译法）
## 语气与风格判定
## 读者理解难点 & 文化注解点
## 修辞隐喻与替换映射
"""
    user_prompt = f"""
源语言：{source_lang}
目标语言：{target_lang}
目标读者：{audience or "普通读者"}
风格模式：{style or "auto"}

原文：
{text[:4000]}
"""
    return chat(sys_prompt, user_prompt, temperature=0.2)


def step2_build_prompt(analysis, source_lang, target_lang, audience, style):
    if style == "auto" or not style.strip():
        style_block = f"""The source text's tone is extracted from analysis above.
Reproduce this tone faithfully in {target_lang}. Match register, rhythm, personality exactly."""
    else:
        style_block = f"""{style}
Apply this style consistently across full text."""

    prompt_full = f"""You are a professional translator. Translate from {source_lang} to {target_lang}.
Think and reason entirely in {target_lang}.

## Target Audience
{audience or "general readers"}

## Translation Style
{style_block}

## Content Background
{analysis}

## Translation Principles
- Complete every sentence, no omission / summarization
- Accuracy first, meaning over literal word
- Keep markdown format fully preserved
- Cultural terms add brief explanation in parentheses
- Natural target language expression, adjust sentence structure freely
"""
    return prompt_full


def step3_draft(source_text, prompt, source_lang, target_lang):
    chunks = split_chunks(source_text)
    subagent_cmd = """Follow all rules in the translation context above.
Translate this chunk COMPLETELY, every sentence, no omission, no summarization.
Keep paragraph count similar, preserve format fully. Only output pure translation result.
"""
    if len(chunks) == 1:
        return chat(prompt + "\n" + subagent_cmd, source_text)
    results = []
    for chunk in chunks:
        t = chat(prompt + "\n" + subagent_cmd, chunk)
        results.append(t)
    return "\n\n".join(results)


def step4_critique(source_text, draft, analysis, source_lang, target_lang):
    return chat(
        "你是严格的翻译审校专家。",
        f"""审校{source_lang}→{target_lang}译文，输出诊断报告：
## 准确性与完整性
## 翻译腔问题（原句 → 建议）
## 修辞与情感保真
## 表达与逻辑
## 总结
只列出问题，不要自行改写译文原文。

原文（前3000字）：\n{source_text[:3000]}
译文：\n{draft[:4000]}
分析：\n{analysis[:1000]}"""
    )


def step5_final(draft, critique, target_lang):
    return chat(
        f"你是{target_lang}母语精修专家，严格依据审校报告逐条修正",
        f"""原文分析、初译草稿、审校问题全部参考上方。
逐条修正术语、表达、不通顺、翻译腔；保持原意不变。
只输出最终纯净定稿译文：
初译：{draft}
审校：{critique}"""
    )


# ─── 翻译线程 ────────────────────────────────────────────
class TranslateWorker(QThread):
    progress = pyqtSignal(str, int)
    analysis_done = pyqtSignal(str)
    critique_done = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, source_text, source_lang, target_lang,
                 style, audience):
        super().__init__()
        self.source_text = source_text
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.style = style
        self.audience = audience

    def run(self):
        try:
            self.progress.emit("📋 第一步：深度分析...", 1)
            analysis = step1_analyze(self.source_text, self.source_lang,
                                     self.target_lang, self.audience, self.style)
            self.analysis_done.emit(analysis)

            self.progress.emit("📝 第二步：组装提示...", 2)
            prompt = step2_build_prompt(analysis, self.source_lang,
                                        self.target_lang, self.audience, self.style)

            self.progress.emit("✍️  第三步：初译...", 3)
            draft = step3_draft(self.source_text, prompt,
                                self.source_lang, self.target_lang)

            self.progress.emit("🔍 第四步：审校...", 4)
            critique = step4_critique(self.source_text, draft, analysis,
                                      self.source_lang, self.target_lang)
            self.critique_done.emit(critique)

            self.progress.emit("✅ 第五步：终稿润色...", 5)
            final = step5_final(draft, critique, self.target_lang)
            self.finished.emit(final)

        except Exception as e:
            self.error.emit(str(e))


# ─── 导出对话框 ───────────────────────────────────────────
class ExportDialog(QDialog):
    def __init__(self, parent, file_type):
        super().__init__(parent)
        self.setWindowTitle("导出设置")
        self.setMinimumWidth(320)

        self.setStyleSheet("""
            QDialog { background-color: #fff; color: #000; }
            QLabel { font-size:14px; color:#000; }
            QRadioButton { font-size:14px; color:#000; padding:4px 0; }
            QPushButton { padding:8px 14px; font-size:14px; border-radius:6px; color:#000; }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20,20,20,20)

        layout.addWidget(QLabel("导出格式："))
        self.fmt_group = QButtonGroup()
        fmt_options = {
            "pdf": "PDF",
            "docx": "Word (.docx)",
            "doc": "Word (.doc)",
            "xlsx": "Excel (.xlsx)",
            "xls": "Excel (.xls)",
            "pptx": "PowerPoint (.pptx)",
            "ppt": "PowerPoint (.ppt)",
        }
        default_fmt = file_type if file_type in fmt_options else "docx"
        for fmt, label in fmt_options.items():
            rb = QRadioButton(label)
            rb.setProperty("val", fmt)
            if fmt == default_fmt:
                rb.setChecked(True)
            self.fmt_group.addButton(rb)
            layout.addWidget(rb)

        layout.addWidget(QLabel("导出模式："))
        self.mode_group = QButtonGroup()
        for mode, label in [
            ("bilingual", "行对照模式（原文+译文）"),
            ("paragraph", "段落对照模式（推荐）"),
            ("translation", "仅译文模式"),
        ]:
            rb = QRadioButton(label)
            rb.setProperty("val", mode)
            if mode == "paragraph":
                rb.setChecked(True)
            self.mode_group.addButton(rb)
            layout.addWidget(rb)

        btn = QPushButton("确认导出")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)

    def get_fmt(self):
        for b in self.fmt_group.buttons():
            if b.isChecked():
                return b.property("val")

    def get_mode(self):
        for b in self.mode_group.buttons():
            if b.isChecked():
                return b.property("val")


# ─── 主窗口 ──────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Translation Agent")
        self.setMinimumSize(980, 780)
        self.file_path = None
        self.file_type = None
        self.file_extra = None
        self.last_result = ""
        self.last_source = ""
        self.last_analysis = ""
        self.last_critique = ""
        self.setup_ui()
        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #FFFAF0; }
            QGroupBox { background:#fff; border:none; border-radius:12px; padding:14px; margin-top:10px; }
            QGroupBox::title { color:#222; font-size:14px; font-weight:bold; }
            QLineEdit, QTextEdit { background:#fff; border:1px solid #EAEAEA; border-radius:8px; padding:9px 12px; font-size:14px; color:#222; }
            QComboBox { background:#fff; border:1px solid #EAEAEA; border-radius:8px; padding:9px 12px; font-size:14px; color:#222; }
            QComboBox QAbstractItemView { background-color:#fff; color:#222; border:1px solid #EAEAEA; }
            QComboBox::item { color:#222; }
            QTabBar::tab { background:#F5F5F5; border-radius:8px 8px 0 0; padding:9px 16px; margin-right:4px; color:#222; }
            QTabBar::tab:selected { background:#D81F26; color:#fff; }
            QPushButton { background:#fff; border:1px solid #EAEAEA; border-radius:8px; padding:8px 16px; font-size:14px; color:#222; }
            QPushButton#primaryBtn { background:#D81F26; color:#fff; border:none; font-weight:bold; }
            QProgressBar { background:#F0F0F0; border-radius:3px; height:6px; }
            QProgressBar::chunk { background:#D81F26; }
        """)

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(16,16,16,16)

        title = QLabel("Translation Agent")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        title.setStyleSheet("color: #222;")
        layout.addWidget(title)

        sub = QLabel("支持 PDF / Word / Excel / PPT · 分析→提示→初译→审校→终稿")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet("color:gray;font-size:12px;")
        layout.addWidget(sub)

        sg = QGroupBox("⚙️ 翻译设置")
        sl = QHBoxLayout(sg)
        for lbl, items, attr in [
            ("源语言", ["English","Chinese","Japanese"], "source_lang"),
            ("目标语言", ["Chinese","English","Japanese"], "target_lang"),
        ]:
            col = QVBoxLayout()
            col.addWidget(QLabel(lbl))
            cb = QComboBox(); cb.addItems(items)
            setattr(self, attr, cb)
            col.addWidget(cb)
            sl.addLayout(col)
        for lbl, ph, attr in [
            ("风格", "formal/conversational/technical/auto", "style_input"),
            ("目标读者", "general/technical/academic/business", "audience_input"),
        ]:
            col = QVBoxLayout()
            col.addWidget(QLabel(lbl))
            le = QLineEdit(); le.setPlaceholderText(ph)
            setattr(self, attr, le)
            col.addWidget(le)
            sl.addLayout(col)
        layout.addWidget(sg)

        ig = QGroupBox("📄 输入内容")
        il = QVBoxLayout(ig)
        self.tabs = QTabWidget()

        tw = QWidget();
        tl = QVBoxLayout(tw)
        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("在此粘贴需要翻译的文字...")
        self.text_input.setMinimumHeight(90)
        # 下面这行是唯一正确、不报错的设置
        from PyQt6.QtGui import QTextOption
        self.text_input.setWordWrapMode(QTextOption.WrapMode.WordWrap)
        tl.addWidget(self.text_input)
        self.tabs.addTab(tw, "📝 粘贴文字")

        uw = QWidget(); ul = QVBoxLayout(uw)
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://example.com/article")
        ul.addWidget(self.url_input); ul.addStretch()
        self.tabs.addTab(uw, "🌐 URL")

        fw = QWidget(); fl = QVBoxLayout(fw)
        fbl = QHBoxLayout()
        self.file_btn = QPushButton("📁 选择文件")
        self.file_btn.clicked.connect(self.choose_file)
        self.file_label = QLabel("支持 .pdf / .docx / .doc / .xlsx / .xls / .pptx / .ppt")
        fbl.addWidget(self.file_btn)
        fbl.addWidget(self.file_label)
        fbl.addStretch()
        fl.addLayout(fbl); fl.addStretch()
        self.tabs.addTab(fw, "📁 上传文件")

        il.addWidget(self.tabs)
        layout.addWidget(ig)

        btn_row = QHBoxLayout()
        self.translate_btn = QPushButton("🚀 开始翻译")
        self.translate_btn.setObjectName("primaryBtn")
        self.translate_btn.clicked.connect(self.do_translate)

        self.export_btn = QPushButton("💾 导出译文")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.do_export)

        self.export_term_btn = QPushButton("📚 导出术语表")
        self.export_term_btn.setEnabled(False)
        self.export_term_btn.clicked.connect(self.export_terms)

        self.export_critique_btn = QPushButton("🔍 导出审校报告")
        self.export_critique_btn.setEnabled(False)
        self.export_critique_btn.clicked.connect(self.export_critique)

        btn_row.addWidget(self.translate_btn, stretch=3)
        btn_row.addWidget(self.export_btn, stretch=1)
        btn_row.addWidget(self.export_term_btn, stretch=1)
        btn_row.addWidget(self.export_critique_btn, stretch=1)
        layout.addLayout(btn_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0,5)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.progress_label)

        rg = QGroupBox("📊 结果")
        rl = QVBoxLayout(rg)
        self.result_tabs = QTabWidget()
        self.analysis_output = QTextEdit()
        self.analysis_output.setReadOnly(True)
        self.result_tabs.addTab(self.analysis_output, "📋 分析报告")
        self.critique_output = QTextEdit()
        self.critique_output.setReadOnly(True)
        self.result_tabs.addTab(self.critique_output, "🔍 审校报告")
        fww = QWidget(); fwl = QVBoxLayout(fww)
        self.result_output = QTextEdit()
        self.result_output.setReadOnly(True)
        fwl.addWidget(self.result_output)
        copy_btn = QPushButton("📋 复制译文")
        copy_btn.clicked.connect(self.copy_result)
        fwl.addWidget(copy_btn)
        self.result_tabs.addTab(fww, "✅ 终稿译文")
        rl.addWidget(self.result_tabs)
        layout.addWidget(rg, stretch=1)

    def choose_file(self):
        try:
            path, _ = QFileDialog.getOpenFileName(
                filter="所有支持格式 (*.pdf *.docx *.doc *.xlsx *.xls *.pptx *.ppt)"
            )
            if not path:
                return

            from file_handler import read_file
            file_type, content, extra = read_file(path)

            self.text_input.setPlainText(content)

            self.file_label.setText(f"已加载：{os.path.basename(path)}")
            self.file_path = path
            self.file_type = file_type
            self.file_extra = extra

        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取失败：{str(e)}")

    def do_translate(self):
        tab = self.tabs.currentIndex()
        source_text = ""
        try:
            if tab == 0:
                source_text = self.text_input.toPlainText().strip()
                if not source_text:
                    QMessageBox.warning(self, "提示", "请输入需要翻译的文字。")
                    return
            elif tab == 1:
                url = self.url_input.text().strip()
                if not url:
                    QMessageBox.warning(self, "提示", "请输入URL。")
                    return
                source_text = fetch_url(url)
            elif tab == 2:
                if not self.file_path:
                    QMessageBox.warning(self, "提示", "请选择文件。")
                    return
                source_text = self.text_input.toPlainText().strip()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            return

        self.last_source = source_text
        self.translate_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.export_term_btn.setEnabled(False)
        self.export_critique_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.analysis_output.clear()
        self.critique_output.clear()
        self.result_output.clear()

        self.worker = TranslateWorker(
            source_text=source_text,
            source_lang=self.source_lang.currentText(),
            target_lang=self.target_lang.currentText(),
            style=self.style_input.text(),
            audience=self.audience_input.text(),
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.analysis_done.connect(self.on_analysis_done)
        self.worker.critique_done.connect(self.on_critique_done)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, msg, step):
        self.progress_label.setText(msg)
        self.progress_bar.setValue(step)

    def on_analysis_done(self, analysis):
        self.last_analysis = analysis
        self.analysis_output.setPlainText(analysis)
        self.export_term_btn.setEnabled(True)

    def on_critique_done(self, critique):
        self.last_critique = critique
        self.critique_output.setPlainText(critique)
        self.export_critique_btn.setEnabled(True)

    def on_finished(self, result):
        self.last_result = result
        self.result_output.setPlainText(result)
        self.progress_label.setText("✅ 翻译完成！")
        self.progress_bar.setValue(5)
        self.translate_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.result_tabs.setCurrentIndex(2)

    def on_error(self, msg):
        QMessageBox.critical(self, "错误", msg)
        self.progress_label.setText("")
        self.translate_btn.setEnabled(True)

    def export_terms(self):
        if not self.last_analysis:
            QMessageBox.warning(self, "提示", "暂无术语表可导出")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出术语表", "术语表.txt",
            "文本文件 (*.txt);;Word文档 (*.docx)")
        if not path:
            return
        try:
            if path.endswith(".docx"):
                from docx import Document
                doc = Document()
                doc.add_heading("术语表", level=1)
                doc.add_paragraph(self.last_analysis)
                doc.save(path)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.last_analysis)
            QMessageBox.information(self, "✅ 导出成功", f"术语表已保存：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    def export_critique(self):
        if not self.last_critique:
            QMessageBox.warning(self, "提示", "暂无审校报告可导出")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出审校报告", "审校报告.txt",
            "文本文件 (*.txt);;Word文档 (*.docx)")
        if not path:
            return
        try:
            if path.endswith(".docx"):
                from docx import Document
                doc = Document()
                doc.add_heading("审校报告", level=1)
                doc.add_paragraph(self.last_critique)
                doc.save(path)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self.last_critique)
            QMessageBox.information(self, "✅ 导出成功", f"审校报告已保存：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    def do_export(self):
        if not self.last_result:
            QMessageBox.warning(self, "提示", "请先完成翻译。")
            return

        dialog = ExportDialog(self, self.file_type or "docx")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        fmt = dialog.get_fmt()
        mode = dialog.get_mode()

        out_path, _ = QFileDialog.getSaveFileName(
            self, "保存文件", f"translation.{fmt}",
            f"{fmt.upper()} 文件 (*.{fmt})"
        )
        if not out_path:
            return

        try:
            from file_handler import (
                export_pdf_bilingual, export_pdf_paragraph, export_pdf_translation,
                export_docx_bilingual, export_docx_paragraph, export_docx_translation,
                export_excel_bilingual, export_excel_translation,
                export_pptx_bilingual, export_pptx_translation,
            )

            if fmt == "pdf":
                if mode == "bilingual":
                    export_pdf_bilingual(self.last_source, self.last_result, out_path)
                elif mode == "paragraph":
                    export_pdf_paragraph(self.last_source, self.last_result, out_path)
                else:
                    export_pdf_translation(self.last_result, out_path)


            elif fmt in ["docx", "doc"]:

                if mode == "bilingual":

                    export_docx_bilingual(self.last_source, self.last_result,

                                          out_path, self.file_path)  # ← 加 self.file_path

                elif mode == "paragraph":

                    export_docx_paragraph(self.last_source, self.last_result,

                                          out_path, self.file_path)  # ← 加 self.file_path

                else:

                    export_docx_translation(self.last_result,

                                            out_path, self.file_path)  # ← 加 self.file_path


            elif fmt in ["xlsx", "xls"]:

                if self.file_extra:

                    if mode == "bilingual":

                        export_excel_bilingual(self.file_extra, self.last_result,

                                               out_path, self.file_path)  # ← 加 self.file_path

                    else:

                        export_excel_translation(self.last_result,

                                                 out_path, self.file_path)  # ← 加 self.file_path

            elif fmt in ["pptx", "ppt"]:
                if self.file_path:
                    if mode == "bilingual":
                        export_pptx_bilingual(self.file_path, self.last_result, out_path)
                    else:
                        export_pptx_translation(self.file_path, self.last_result, out_path)
                else:
                    QMessageBox.warning(self, "提示", "需先导入 PPT 文件")
                    return

            QMessageBox.information(self, "✅ 导出成功", f"已保存：\n{out_path}")

        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    def copy_result(self):
        text = self.result_output.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "✅", "已复制译文")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())