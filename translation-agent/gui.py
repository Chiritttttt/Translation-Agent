import os
import sys
from dotenv import load_dotenv
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QTabWidget, QMessageBox, QGroupBox, QProgressBar, QDialog,
    QRadioButton, QButtonGroup, QFrame, QGraphicsDropShadowEffect, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QColor, QPalette, QIcon, QPixmap
import openai
import requests
from bs4 import BeautifulSoup

load_dotenv()

client = openai.OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url=os.getenv("OPENAI_BASE_URL"),
)
model = os.getenv("OPENAI_MODEL", "deepseek-chat")


# ═══════════════════════════════════════════════════════════
# Arco Design 颜色系统
# ═══════════════════════════════════════════════════════════
class ArcoColors:
    """Arco Design 色板"""

    # ── 主色 ──
    PRIMARY          = "#165DFF"
    PRIMARY_LIGHT    = "#E8F3FF"
    PRIMARY_LIGHT2   = "#F2F7FF"
    PRIMARY_HOVER    = "#4080FF"
    PRIMARY_ACTIVE   = "#0E42D2"

    # ── 功能色 ──
    SUCCESS          = "#00B42A"
    SUCCESS_LIGHT    = "#E8FFEA"
    WARNING          = "#FF7D00"
    WARNING_LIGHT    = "#FFF7E8"
    DANGER           = "#F53F3F"
    DANGER_LIGHT     = "#FFECE8"
    INFO             = "#165DFF"
    INFO_LIGHT       = "#E8F3FF"

    # ── 中性色 ──
    TEXT_PRIMARY      = "#1D2129"
    TEXT_REGULAR      = "#4E5969"
    TEXT_SECONDARY    = "#86909C"
    TEXT_DISABLED     = "#C9CDD4"
    TEXT_CAPTION      = "#C9CDD4"

    BORDER           = "#E5E6EB"
    BORDER_LIGHT     = "#F2F3F5"
    FILL1            = "#F7F8FA"
    FILL2            = "#F2F3F5"
    FILL3            = "#E5E6EB"
    FILL4            = "#C9CDD4"

    BG_PAGE          = "#F2F3F5"
    BG_CARD          = "#FFFFFF"
    BG_ELEVATION     = "#FFFFFF"

    # ── 阴影 ──
    SHADOW_SM        = "0 1px 2px 0 rgba(0,0,0,0.03), 0 1px 6px -1px rgba(0,0,0,0.02), 0 2px 4px 0 rgba(0,0,0,0.02)"
    SHADOW_MD        = "0 3px 6px -4px rgba(0,0,0,0.12), 0 6px 16px 0 rgba(0,0,0,0.08), 0 9px 28px 8px rgba(0,0,0,0.05)"
    SHADOW_LG        = "0 6px 30px 5px rgba(0,0,0,0.05), 0 16px 24px 2px rgba(0,0,0,0.04), 0 8px 10px -5px rgba(0,0,0,0.08)"

    # ── 圆角 ──
    RADIUS_SM = "4px"
    RADIUS_MD = "8px"
    RADIUS_LG = "12px"

    # ── 字体 ──
    FONT_FAMILY = "'Inter', 'SF Pro Display', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans SC', sans-serif"


# ═══════════════════════════════════════════════════════════
# Arco Design QSS 样式表
# ═══════════════════════════════════════════════════════════
def arco_global_stylesheet():
    C = ArcoColors
    return f"""
    /* ── 全局 ── */
    QWidget {{
        font-family: {C.FONT_FAMILY};
        font-size: 14px;
        color: {C.TEXT_PRIMARY};
    }}

    /* ── 主窗口 ── */
    QMainWindow {{
        background-color: {C.BG_PAGE};
    }}

    /* ── 卡片容器 (QGroupBox) ── */
    QGroupBox {{
        background: {C.BG_CARD};
        border: 1px solid {C.BORDER_LIGHT};
        border-radius: {C.RADIUS_LG};
        padding: 20px 20px 16px 20px;
        margin-top: 0px;
        margin-bottom: 4px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        color: {C.TEXT_PRIMARY};
        font-size: 14px;
        font-weight: 600;
        padding: 0 0 12px 0;
        left: 20px;
    }}

    /* ── 输入框 ── */
    QLineEdit {{
        background: {C.BG_CARD};
        border: 1px solid {C.BORDER};
        border-radius: {C.RADIUS_MD};
        padding: 8px 14px;
        font-size: 14px;
        color: {C.TEXT_PRIMARY};
        selection-background-color: {C.PRIMARY_LIGHT};
        selection-color: {C.PRIMARY};
    }}
    QLineEdit:hover {{
        border-color: {C.PRIMARY_HOVER};
    }}
    QLineEdit:focus {{
        border-color: {C.PRIMARY};
        border-width: 2px;
        padding: 7px 13px;
    }}
    QLineEdit:disabled {{
        background: {C.FILL1};
        color: {C.TEXT_DISABLED};
        border-color: {C.BORDER_LIGHT};
    }}

    /* ── 文本编辑框 ── */
    QTextEdit {{
        background: {C.BG_CARD};
        border: 1px solid {C.BORDER};
        border-radius: {C.RADIUS_MD};
        padding: 10px 14px;
        font-size: 14px;
        line-height: 1.6;
        color: {C.TEXT_PRIMARY};
        selection-background-color: {C.PRIMARY_LIGHT};
        selection-color: {C.PRIMARY};
    }}
    QTextEdit:hover {{
        border-color: {C.PRIMARY_HOVER};
    }}
    QTextEdit:focus {{
        border-color: {C.PRIMARY};
        border-width: 2px;
        padding: 9px 13px;
    }}
    QTextEdit:read-only {{
        background: {C.FILL1};
        color: {C.TEXT_REGULAR};
        border-color: {C.BORDER_LIGHT};
    }}

    /* ── 下拉框 ── */
    QComboBox {{
        background: {C.BG_CARD};
        border: 1px solid {C.BORDER};
        border-radius: {C.RADIUS_MD};
        padding: 8px 36px 8px 14px;
        font-size: 14px;
        color: {C.TEXT_PRIMARY};
        min-height: 22px;
    }}
    QComboBox:hover {{
        border-color: {C.PRIMARY_HOVER};
    }}
    QComboBox:focus, QComboBox::drop-down:pressed {{
        border-color: {C.PRIMARY};
    }}
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: right center;
        width: 28px;
        border: none;
        background: transparent;
    }}
    QComboBox::down-arrow {{
        width: 12px;
        height: 12px;
        image: none;
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 6px solid {C.TEXT_SECONDARY};
    }}
    QComboBox QAbstractItemView {{
        background: {C.BG_ELEVATION};
        color: {C.TEXT_PRIMARY};
        border: 1px solid {C.BORDER};
        border-radius: {C.RADIUS_MD};
        padding: 4px 0;
        selection-background-color: {C.PRIMARY_LIGHT};
        selection-color: {C.PRIMARY};
        outline: none;
    }}
    QComboBox QAbstractItemView::item {{
        height: 36px;
        padding: 0 14px;
        color: {C.TEXT_PRIMARY};
    }}
    QComboBox QAbstractItemView::item:hover {{
        background-color: {C.FILL1};
    }}
    QComboBox QAbstractItemView::item:selected {{
        background-color: {C.PRIMARY_LIGHT};
        color: {C.PRIMARY};
    }}

    /* ── 标签页 ── */
    QTabWidget::pane {{
        border: none;
        background: transparent;
        border-radius: {C.RADIUS_MD};
    }}
    QTabBar {{
        background: transparent;
    }}
    QTabBar::tab {{
        background: transparent;
        color: {C.TEXT_SECONDARY};
        font-size: 14px;
        font-weight: 500;
        padding: 10px 20px;
        margin-right: 2px;
        border: none;
        border-radius: {C.RADIUS_MD} {C.RADIUS_MD} 0 0;
        border-bottom: 2px solid transparent;
    }}
    QTabBar::tab:hover {{
        color: {C.TEXT_PRIMARY};
        background: {C.FILL1};
    }}
    QTabBar::tab:selected {{
        color: {C.PRIMARY};
        background: {C.BG_CARD};
        border-bottom: 2px solid {C.PRIMARY};
        font-weight: 600;
    }}

    /* ── 主按钮 (Primary) ── */
    QPushButton#primaryBtn {{
        background: {C.PRIMARY};
        color: #FFFFFF;
        border: none;
        border-radius: {C.RADIUS_MD};
        padding: 10px 24px;
        font-size: 15px;
        font-weight: 600;
    }}
    QPushButton#primaryBtn:hover {{
        background: {C.PRIMARY_HOVER};
    }}
    QPushButton#primaryBtn:pressed {{
        background: {C.PRIMARY_ACTIVE};
    }}
    QPushButton#primaryBtn:disabled {{
        background: {C.FILL3};
        color: {C.TEXT_DISABLED};
    }}

    /* ── 次按钮 (Secondary/Outline) ── */
    QPushButton#outlineBtn {{
        background: transparent;
        color: {C.TEXT_REGULAR};
        border: 1px solid {C.BORDER};
        border-radius: {C.RADIUS_MD};
        padding: 8px 16px;
        font-size: 13px;
        font-weight: 500;
    }}
    QPushButton#outlineBtn:hover {{
        color: {C.PRIMARY};
        border-color: {C.PRIMARY};
        background: {C.PRIMARY_LIGHT2};
    }}
    QPushButton#outlineBtn:pressed {{
        color: {C.PRIMARY_ACTIVE};
        border-color: {C.PRIMARY_ACTIVE};
    }}
    QPushButton#outlineBtn:disabled {{
        color: {C.TEXT_DISABLED};
        border-color: {C.BORDER_LIGHT};
    }}

    /* ── 文字按钮 (Text) ── */
    QPushButton#textBtn {{
        background: transparent;
        color: {C.TEXT_REGULAR};
        border: none;
        border-radius: {C.RADIUS_SM};
        padding: 8px 12px;
        font-size: 13px;
    }}
    QPushButton#textBtn:hover {{
        background: {C.FILL1};
        color: {C.PRIMARY};
    }}
    QPushButton#textBtn:disabled {{
        color: {C.TEXT_DISABLED};
    }}

    /* ── 普通按钮 (兜底，无 objectName 的) ── */
    QPushButton {{
        background: transparent;
        color: {C.TEXT_REGULAR};
        border: 1px solid {C.BORDER};
        border-radius: {C.RADIUS_MD};
        padding: 8px 16px;
        font-size: 13px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        color: {C.PRIMARY};
        border-color: {C.PRIMARY};
        background: {C.PRIMARY_LIGHT2};
    }}
    QPushButton:pressed {{
        color: {C.PRIMARY_ACTIVE};
        border-color: {C.PRIMARY_ACTIVE};
    }}
    QPushButton:disabled {{
        color: {C.TEXT_DISABLED};
        border-color: {C.BORDER_LIGHT};
    }}

    /* ── 进度条 ── */
    QProgressBar {{
        background: {C.FILL2};
        border: none;
        border-radius: 100px;
        height: 8px;
        text-align: center;
    }}
    QProgressBar::chunk {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 {C.PRIMARY}, stop:1 {C.PRIMARY_HOVER});
        border-radius: 100px;
    }}

    /* ── 单选按钮 ── */
    QRadioButton {{
        font-size: 14px;
        color: {C.TEXT_PRIMARY};
        spacing: 8px;
        padding: 6px 0;
    }}
    QRadioButton::indicator {{
        width: 16px;
        height: 16px;
        border: 2px solid {C.BORDER};
        border-radius: 50%;
        background: {C.BG_CARD};
    }}
    QRadioButton::indicator:hover {{
        border-color: {C.PRIMARY_HOVER};
    }}
    QRadioButton::indicator:checked {{
        border: 5px solid {C.PRIMARY};
        background: transparent;
    }}

    /* ── 表格 ── */
    QTableWidget {{
        background: {C.BG_CARD};
        alternate-background-color: {C.FILL1};
        border: 1px solid {C.BORDER_LIGHT};
        border-radius: {C.RADIUS_MD};
        font-size: 13px;
        gridline-color: {C.BORDER_LIGHT};
        selection-background-color: {C.PRIMARY_LIGHT};
        selection-color: {C.PRIMARY};
    }}
    QTableWidget::item {{
        padding: 10px 14px;
        border-bottom: 1px solid {C.BORDER_LIGHT};
    }}
    QTableWidget::item:hover {{
        background: {C.FILL1};
    }}
    QHeaderView::section {{
        background: {C.FILL1};
        color: {C.TEXT_REGULAR};
        padding: 10px 14px;
        border: none;
        border-bottom: 1px solid {C.BORDER};
        border-right: 1px solid {C.BORDER_LIGHT};
        font-weight: 600;
        font-size: 13px;
    }}
    QTableWidget QTableWidget {{
        border-radius: 0px;
    }}
    """


# ═══════════════════════════════════════════════════════════
# 辅助函数：创建分隔线
# ═══════════════════════════════════════════════════════════
def make_h_line(color=None):
    """水平分隔线"""
    C = ArcoColors
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setFixedHeight(1)
    line.setStyleSheet(f"background-color: {color or C.BORDER_LIGHT}; border: none; margin: 4px 0;")
    return line


def make_shadow_widget(widget, shadow_type="sm"):
    """给 widget 添加 Arco Design 风格阴影"""
    C = ArcoColors
    shadow = QGraphicsDropShadowEffect()
    shadow.setBlurRadius(12)
    shadow.setOffset(0, 2)
    shadow.setColor(QColor(0, 0, 0, 20))
    if shadow_type == "md":
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 30))
    widget.setGraphicsEffect(shadow)


# ═══════════════════════════════════════════════════════════
# AI 调用
# ═══════════════════════════════════════════════════════════
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


# ═══════════════════════════════════════════════════════════
# 五步翻译工作流
# ═══════════════════════════════════════════════════════════
def step1_analyze(text, source_lang, target_lang, audience, style):
    sys_prompt = """你是专业翻译前置分析师。
输出严格按照以下固定Markdown结构：
## 概要
## 术语表（原文 → 译法）
请逐行列出所有专业术语、缩写、惯用表达、专有名词。
格式要求：每行一条，用 → 连接，例如：
- machine learning → 机器学习
- large language model → 大语言模型
- the state of the art → 最先进的
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
    from glossary import get_glossary_prompt_block

    if style == "auto" or not style.strip():
        style_block = f"""The source text's tone is extracted from analysis above.
Reproduce this tone faithfully in {target_lang}. Match register, rhythm, personality exactly."""
    else:
        style_block = f"""{style}
Apply this style consistently across full text."""

    glossary_block = get_glossary_prompt_block(source_lang, target_lang)

    prompt_full = f"""You are a professional translator. Translate from {source_lang} to {target_lang}.
Think and reason entirely in {target_lang}.

## Target Audience
{audience or "general readers"}

## Translation Style
{style_block}

{glossary_block}

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


# ═══════════════════════════════════════════════════════════
# 翻译线程
# ═══════════════════════════════════════════════════════════
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
            self.progress.emit("第一步：深度分析...", 1)
            analysis = step1_analyze(self.source_text, self.source_lang,
                                     self.target_lang, self.audience, self.style)
            self.analysis_done.emit(analysis)

            from glossary import extract_terms_from_analysis, add_terms_batch
            new_terms = extract_terms_from_analysis(analysis, self.source_lang, self.target_lang)
            if new_terms:
                added, _ = add_terms_batch(new_terms, self.source_lang, self.target_lang)
                self.progress.emit(f"术语入库：新增 {added} 条", 1)

            self.progress.emit("第二步：组装提示...", 2)
            prompt = step2_build_prompt(analysis, self.source_lang,
                                        self.target_lang, self.audience, self.style)

            self.progress.emit("第三步：初译...", 3)
            draft = step3_draft(self.source_text, prompt,
                                self.source_lang, self.target_lang)

            self.progress.emit("第四步：审校...", 4)
            critique = step4_critique(self.source_text, draft, analysis,
                                      self.source_lang, self.target_lang)
            self.critique_done.emit(critique)

            self.progress.emit("第五步：终稿润色...", 5)
            final = step5_final(draft, critique, self.target_lang)
            self.finished.emit(final)

        except Exception as e:
            self.error.emit(str(e))


# ═══════════════════════════════════════════════════════════
# 导出对话框 — Arco Design 风格
# ═══════════════════════════════════════════════════════════
class ExportDialog(QDialog):
    def __init__(self, parent, file_type):
        super().__init__(parent)
        self.setWindowTitle("导出设置")
        self.setMinimumWidth(360)
        self.setMinimumHeight(420)
        C = ArcoColors
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {C.BG_PAGE};
                color: {C.TEXT_PRIMARY};
            }}
            QLabel {{
                font-size: 14px;
                color: {C.TEXT_PRIMARY};
                font-weight: 600;
            }}
            QRadioButton {{
                font-size: 14px;
                color: {C.TEXT_PRIMARY};
                spacing: 8px;
                padding: 7px 12px;
                border-radius: {C.RADIUS_MD};
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER_LIGHT};
            }}
            QRadioButton:hover {{
                border-color: {C.PRIMARY_HOVER};
                background: {C.PRIMARY_LIGHT2};
            }}
            QRadioButton::indicator {{
                width: 16px;
                height: 16px;
                border: 2px solid {C.BORDER};
                border-radius: 50%;
                background: {C.BG_CARD};
            }}
            QRadioButton::indicator:checked {{
                border: 5px solid {C.PRIMARY};
            }}
            QPushButton {{
                background: {C.PRIMARY};
                color: #FFFFFF;
                border: none;
                border-radius: {C.RADIUS_MD};
                padding: 10px 24px;
                font-size: 14px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: {C.PRIMARY_HOVER};
            }}
        """)

        # 带阴影的卡片布局
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_LG};
                padding: 24px;
            }}
        """)
        make_shadow_widget(card, "md")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(16)
        card_layout.setContentsMargins(24, 24, 24, 24)

        # 外层布局
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.addWidget(card)

        card_layout.addWidget(QLabel("导出格式"))

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
        fmt_grid = QVBoxLayout()
        fmt_grid.setSpacing(6)
        for fmt, label in fmt_options.items():
            rb = QRadioButton(label)
            rb.setProperty("val", fmt)
            if fmt == default_fmt:
                rb.setChecked(True)
            self.fmt_group.addButton(rb)
            fmt_grid.addWidget(rb)
        card_layout.addLayout(fmt_grid)

        card_layout.addWidget(make_h_line())
        card_layout.addWidget(QLabel("导出模式"))

        self.mode_group = QButtonGroup()
        mode_grid = QVBoxLayout()
        mode_grid.setSpacing(6)
        for mode, label in [
            ("paragraph", "段落对照模式（推荐）"),
            ("bilingual", "行对照模式（原文+译文）"),
            ("translation", "仅译文模式"),
        ]:
            rb = QRadioButton(label)
            rb.setProperty("val", mode)
            if mode == "paragraph":
                rb.setChecked(True)
            self.mode_group.addButton(rb)
            mode_grid.addWidget(rb)
        card_layout.addLayout(mode_grid)

        card_layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn = QPushButton("确认导出")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(self.accept)
        btn_row.addWidget(btn)
        card_layout.addLayout(btn_row)

    def get_fmt(self):
        for b in self.fmt_group.buttons():
            if b.isChecked():
                return b.property("val")

    def get_mode(self):
        for b in self.mode_group.buttons():
            if b.isChecked():
                return b.property("val")


# ═══════════════════════════════════════════════════════════
# 主窗口 — Arco Design 风格
# ═══════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Translation Agent")
        self.setMinimumSize(1060, 820)
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
        self.setStyleSheet(arco_global_stylesheet())

    def setup_ui(self):
        C = ArcoColors

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 20, 24, 20)

        # ── 顶部标题栏 ──
        header = QFrame()
        header.setStyleSheet("background: transparent; border: none;")
        header_layout = QVBoxLayout(header)
        header_layout.setSpacing(4)
        header_layout.setContentsMargins(0, 0, 0, 0)

        title = QLabel("Translation Agent")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont(C.FONT_FAMILY.split(",")[0].strip().strip("'"), 22, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {C.TEXT_PRIMARY}; background: transparent; border: none;")
        header_layout.addWidget(title)

        sub = QLabel("支持 PDF / Word / Excel / PPT  ·  分析 → 提示 → 初译 → 审校 → 终稿")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; background: transparent; border: none; letter-spacing: 0.5px;")
        header_layout.addWidget(sub)

        layout.addWidget(header)

        # ── 翻译设置卡片 ──
        sg = QFrame()
        sg.setObjectName("card")
        sg.setStyleSheet(f"""
            QFrame#card {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_LG};
            }}
        """)
        make_shadow_widget(sg, "sm")
        sl = QHBoxLayout(sg)
        sl.setSpacing(24)
        sl.setContentsMargins(20, 16, 20, 16)

        # 卡片小标题（放在卡片上方）
        settings_title = QLabel("翻译设置")
        settings_title.setFont(QFont(C.FONT_FAMILY.split(",")[0].strip().strip("'"), 13, QFont.Weight.Bold))
        settings_title.setStyleSheet(f"color: {C.TEXT_PRIMARY}; border: none; background: transparent;")

        layout.addWidget(settings_title)
        layout.addWidget(sg)

        for lbl_text, items, attr in [
            ("源语言", ["English","Chinese","Japanese"], "source_lang"),
            ("目标语言", ["Chinese","English","Japanese"], "target_lang"),
        ]:
            col = QVBoxLayout()
            col.setSpacing(6)
            lbl = QLabel(lbl_text)
            lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; font-weight: 500; border: none; background: transparent;")
            lbl.setFixedHeight(18)
            col.addWidget(lbl)
            cb = QComboBox()
            cb.addItems(items)
            cb.setFixedHeight(38)
            cb.setMinimumWidth(130)
            setattr(self, attr, cb)
            col.addWidget(cb)
            sl.addLayout(col)

        for lbl_text, ph, attr in [
            ("风格", "formal / conversational / technical / auto", "style_input"),
            ("目标读者", "general / technical / academic / business", "audience_input"),
        ]:
            col = QVBoxLayout()
            col.setSpacing(6)
            lbl = QLabel(lbl_text)
            lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; font-weight: 500; border: none; background: transparent;")
            lbl.setFixedHeight(18)
            col.addWidget(lbl)
            le = QLineEdit()
            le.setPlaceholderText(ph)
            le.setFixedHeight(38)
            le.setMinimumWidth(130)
            setattr(self, attr, le)
            col.addWidget(le)
            sl.addLayout(col)

        # ── 输入内容卡片 ──
        ig = QFrame()
        ig.setObjectName("inputCard")
        ig.setStyleSheet(f"""
            QFrame#inputCard {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_LG};
            }}
        """)
        make_shadow_widget(ig, "sm")
        il = QVBoxLayout(ig)
        il.setSpacing(8)
        il.setContentsMargins(20, 16, 20, 16)

        input_title = QLabel("输入内容")
        input_title.setFont(QFont(C.FONT_FAMILY.split(",")[0].strip().strip("'"), 13, QFont.Weight.Bold))
        input_title.setStyleSheet(f"color: {C.TEXT_PRIMARY}; border: none; background: transparent;")
        il.addWidget(input_title)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_MD};
                background: {C.BG_CARD};
                padding: 4px;
            }}
        """)

        # Tab 1: 粘贴文字
        tw = QWidget()
        tl = QVBoxLayout(tw)
        tl.setContentsMargins(8, 8, 8, 8)
        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("在此粘贴需要翻译的文字...")
        self.text_input.setMinimumHeight(100)
        from PyQt6.QtGui import QTextOption
        self.text_input.setWordWrapMode(QTextOption.WrapMode.WordWrap)
        tl.addWidget(self.text_input)
        self.tabs.addTab(tw, "  粘贴文字  ")

        # Tab 2: URL
        uw = QWidget()
        ul = QVBoxLayout(uw)
        ul.setContentsMargins(8, 12, 8, 8)
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://example.com/article")
        self.url_input.setFixedHeight(38)
        ul.addWidget(self.url_input)
        ul.addStretch()
        self.tabs.addTab(uw, "  URL  ")

        # Tab 3: 上传文件
        fw = QWidget()
        fl = QVBoxLayout(fw)
        fl.setContentsMargins(8, 12, 8, 8)
        fbl = QHBoxLayout()
        self.file_btn = QPushButton("选择文件")
        self.file_btn.setObjectName("outlineBtn")
        self.file_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.file_btn.clicked.connect(self.choose_file)
        self.file_label = QLabel("支持 .pdf / .docx / .doc / .xlsx / .xls / .pptx / .ppt")
        self.file_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; border: none; background: transparent;")
        fbl.addWidget(self.file_btn)
        fbl.addWidget(self.file_label)
        fbl.addStretch()
        fl.addLayout(fbl)
        fl.addStretch()
        self.tabs.addTab(fw, "  上传文件  ")

        il.addWidget(self.tabs)

        layout.addWidget(ig)

        # ── 操作栏 ──
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.translate_btn = QPushButton("  开始翻译  ")
        self.translate_btn.setObjectName("primaryBtn")
        self.translate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.translate_btn.setMinimumHeight(40)
        self.translate_btn.clicked.connect(self.do_translate)

        self.export_btn = QPushButton("  导出译文  ")
        self.export_btn.setObjectName("outlineBtn")
        self.export_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.export_btn.setMinimumHeight(36)
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.do_export)

        self.export_term_btn = QPushButton("  导出术语  ")
        self.export_term_btn.setObjectName("outlineBtn")
        self.export_term_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.export_term_btn.setMinimumHeight(36)
        self.export_term_btn.clicked.connect(self.export_glossary)

        self.export_critique_btn = QPushButton("  导出审校  ")
        self.export_critique_btn.setObjectName("outlineBtn")
        self.export_critique_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.export_critique_btn.setMinimumHeight(36)
        self.export_critique_btn.setEnabled(False)
        self.export_critique_btn.clicked.connect(self.export_critique)

        self.glossary_btn = QPushButton("  术语库  ")
        self.glossary_btn.setObjectName("outlineBtn")
        self.glossary_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.glossary_btn.setMinimumHeight(36)
        self.glossary_btn.clicked.connect(self.open_glossary_manager)

        btn_row.addWidget(self.translate_btn, stretch=0)
        btn_row.addSpacing(12)
        btn_row.addWidget(self.export_btn)
        btn_row.addWidget(self.export_term_btn)
        btn_row.addWidget(self.export_critique_btn)
        btn_row.addWidget(self.glossary_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # ── 进度条 ──
        progress_container = QFrame()
        progress_container.setStyleSheet("background: transparent; border: none;")
        pc_layout = QVBoxLayout(progress_container)
        pc_layout.setSpacing(6)
        pc_layout.setContentsMargins(0, 0, 0, 0)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 5)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        pc_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; background: transparent; border: none;")
        pc_layout.addWidget(self.progress_label)
        layout.addWidget(progress_container)

        # ── 结果卡片 ──
        rg = QFrame()
        rg.setObjectName("resultCard")
        rg.setStyleSheet(f"""
            QFrame#resultCard {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_LG};
            }}
        """)
        make_shadow_widget(rg, "sm")
        rl = QVBoxLayout(rg)
        rl.setSpacing(8)
        rl.setContentsMargins(20, 16, 20, 16)

        result_title = QLabel("翻译结果")
        result_title.setFont(QFont(C.FONT_FAMILY.split(",")[0].strip().strip("'"), 13, QFont.Weight.Bold))
        result_title.setStyleSheet(f"color: {C.TEXT_PRIMARY}; border: none; background: transparent;")
        rl.addWidget(result_title)

        self.result_tabs = QTabWidget()
        self.result_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_MD};
                background: {C.BG_CARD};
                padding: 4px;
            }}
        """)

        self.analysis_output = QTextEdit()
        self.analysis_output.setReadOnly(True)
        self.result_tabs.addTab(self.analysis_output, "  分析报告  ")

        self.critique_output = QTextEdit()
        self.critique_output.setReadOnly(True)
        self.result_tabs.addTab(self.critique_output, "  审校报告  ")

        fww = QWidget()
        fwl = QVBoxLayout(fww)
        fwl.setContentsMargins(8, 8, 8, 8)
        self.result_output = QTextEdit()
        self.result_output.setReadOnly(True)
        fwl.addWidget(self.result_output)
        copy_btn = QPushButton("  复制译文  ")
        copy_btn.setObjectName("textBtn")
        copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_btn.clicked.connect(self.copy_result)
        fwl.addWidget(copy_btn)
        self.result_tabs.addTab(fww, "  终稿译文  ")

        rl.addWidget(self.result_tabs)

        layout.addWidget(rg, stretch=1)

    def _ocr_progress_handler(self, current, total, message):
        """OCR 进度回调，在主线程安全地更新 UI"""
        from PyQt6.QtCore import QMetaObject, Q_ARG
        QMetaObject.invokeMethod(
            self.progress_label, "setText",
            Q_ARG(str, message)
        )
        if total > 0:
            QMetaObject.invokeMethod(
                self.progress_bar, "setValue",
                Q_ARG(int, int(current * 5 / max(total, 1)))
            )
        QApplication.processEvents()

    def choose_file(self):
        try:
            path, _ = QFileDialog.getOpenFileName(
                filter="所有支持格式 (*.pdf *.docx *.doc *.xlsx *.xls *.pptx *.ppt)"
            )
            if not path:
                return

            from file_handler import read_file, set_ocr_progress_callback

            ext = path.lower().split(".")[-1]
            if ext == "pdf":
                set_ocr_progress_callback(self._ocr_progress_handler)
                self.progress_label.setText("正在读取 PDF...")
                self.progress_bar.setRange(0, 5)
                self.progress_bar.setValue(0)
                QApplication.processEvents()

            file_type, content, extra = read_file(path)

            set_ocr_progress_callback(None)
            self.progress_bar.setValue(0)
            self.progress_label.setText("")

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
            source_lang=self._get_source_lang(),
            target_lang=self._get_target_lang(),
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
        self.progress_label.setText("翻译完成！")
        self.progress_bar.setValue(5)
        self.translate_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.result_tabs.setCurrentIndex(2)

    def on_error(self, msg):
        QMessageBox.critical(self, "错误", msg)
        self.progress_label.setText("")
        self.translate_btn.setEnabled(True)

    def _get_source_lang(self):
        return self.source_lang.currentText() if hasattr(self, 'source_lang') and self.source_lang else "English"

    def _get_target_lang(self):
        return self.target_lang.currentText() if hasattr(self, 'target_lang') and self.target_lang else "Chinese"

    def export_glossary(self):
        """导出术语库（只包含词汇和表达，不含分析报告全文）"""
        from glossary import export_glossary_text, get_terms
        source_lang = self._get_source_lang()
        target_lang = self._get_target_lang()
        terms = get_terms(source_lang, target_lang)
        if not terms:
            QMessageBox.information(self, "提示",
                f"当前语言对（{source_lang} → {target_lang}）的术语库为空。\n\n"
                "翻译后会自动从分析报告中提取术语入库。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出术语库", f"术语表_{source_lang}_{target_lang}.txt",
            "文本文件 (*.txt);;JSON 文件 (*.json);;Word文档 (*.docx)")
        if not path:
            return
        try:
            if path.endswith(".json"):
                from glossary import export_glossary_json
                content = export_glossary_json(source_lang, target_lang)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
            elif path.endswith(".docx"):
                from docx import Document
                from docx.shared import RGBColor
                doc = Document()
                doc.add_heading(f"术语表 ({source_lang} → {target_lang})", level=1)

                categories = {}
                for t in terms:
                    cat = t.get("category", "术语")
                    if cat not in categories:
                        categories[cat] = []
                    categories[cat].append(t)

                for cat, cat_terms in categories.items():
                    doc.add_heading(f"【{cat}】", level=2)
                    table = doc.add_table(rows=1, cols=3)
                    table.style = "Table Grid"
                    hdr = table.rows[0].cells
                    for idx, txt in enumerate(["原文", "译文", "分类"]):
                        run = hdr[idx].paragraphs[0].add_run(txt)
                        run.bold = True
                        run.font.color.rgb = RGBColor(0x16, 0x5D, 0xFF)
                    for t in cat_terms:
                        row = table.add_row().cells
                        row[0].paragraphs[0].add_run(t["source"])
                        row[1].paragraphs[0].add_run(t["target"])
                        row[2].paragraphs[0].add_run(t.get("category", ""))
                    doc.add_paragraph("")

                doc.save(path)
            else:
                content = export_glossary_text(source_lang, target_lang)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)

            QMessageBox.information(self, "导出成功",
                f"术语库已保存（{len(terms)} 条）：\n{path}")
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
            QMessageBox.information(self, "导出成功", f"审校报告已保存：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    def _get_export_suffix(self):
        """根据源语言/目标语言生成 EC/CE 后缀。
        中译英 = CE (Chinese→English)，英译中 = EC (English→Chinese)
        其他语言对统一用方向缩写。"""
        sl = self.source_lang.currentText().lower()
        tl = self.target_lang.currentText().lower()
        if sl == "chinese" and tl == "english":
            return "CE"
        elif sl == "english" and tl == "chinese":
            return "EC"
        else:
            return f"{sl[:2].upper()}{tl[:2].upper()}"

    def _get_default_export_name(self, fmt):
        """生成默认导出文件名：原文件名-EC/CE-agent.fmt"""
        suffix = self._get_export_suffix()
        if self.file_path and os.path.exists(self.file_path):
            name = os.path.splitext(os.path.basename(self.file_path))[0]
            return f"{name}-{suffix}-agent.{fmt}"
        else:
            return f"translation-{suffix}-agent.{fmt}"

    def do_export(self):
        if not self.last_result:
            QMessageBox.warning(self, "提示", "请先完成翻译。")
            return

        dialog = ExportDialog(self, self.file_type or "docx")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        fmt = dialog.get_fmt()
        mode = dialog.get_mode()
        default_name = self._get_default_export_name(fmt)

        out_path, _ = QFileDialog.getSaveFileName(
            self, "保存文件", default_name,
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
                                          out_path, self.file_path)
                elif mode == "paragraph":
                    export_docx_paragraph(self.last_source, self.last_result,
                                          out_path, self.file_path)
                else:
                    export_docx_translation(self.last_result,
                                            out_path, self.file_path)

            elif fmt in ["xlsx", "xls"]:
                if self.file_extra:
                    if mode == "bilingual":
                        export_excel_bilingual(self.file_extra, self.last_result,
                                               out_path, self.file_path)
                    else:
                        export_excel_translation(self.last_result,
                                                 out_path, self.file_path)

            elif fmt in ["pptx", "ppt"]:
                if self.file_path:
                    if mode == "bilingual":
                        export_pptx_bilingual(self.file_path, self.last_result, out_path)
                    else:
                        export_pptx_translation(self.file_path, self.last_result, out_path)
                else:
                    QMessageBox.warning(self, "提示", "需先导入 PPT 文件")
                    return

            QMessageBox.information(self, "导出成功", f"已保存：\n{out_path}")

        except Exception as e:
            QMessageBox.critical(self, "失败", str(e))

    def copy_result(self):
        text = self.result_output.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "成功", "已复制译文")

    def open_glossary_manager(self):
        """打开术语库管理窗口 — Arco Design 风格"""
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
        from glossary import (
            get_terms, get_all_lang_pairs, delete_term, delete_lang_pair,
            clear_all, import_glossary_from_text, load_glossary,
        )

        C = ArcoColors

        dialog = QDialog(self)
        dialog.setWindowTitle("术语库管理")
        dialog.setMinimumSize(760, 560)
        dialog.setStyleSheet(f"""
            QDialog {{
                background-color: {C.BG_PAGE};
                color: {C.TEXT_PRIMARY};
            }}
            QLabel {{
                font-size: 13px;
                color: {C.TEXT_PRIMARY};
                border: none; background: transparent;
            }}
            QPushButton {{
                background: transparent;
                color: {C.TEXT_REGULAR};
                border: 1px solid {C.BORDER};
                border-radius: {C.RADIUS_MD};
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                color: {C.PRIMARY};
                border-color: {C.PRIMARY};
                background: {C.PRIMARY_LIGHT2};
            }}
            QPushButton:disabled {{
                color: {C.TEXT_DISABLED};
                border-color: {C.BORDER_LIGHT};
            }}
            QPushButton#dangerBtn {{
                color: {C.DANGER};
                border-color: {C.DANGER};
            }}
            QPushButton#dangerBtn:hover {{
                background: {C.DANGER_LIGHT};
                border-color: {C.DANGER};
                color: {C.DANGER};
            }}
            QPushButton#primarySmall {{
                background: {C.PRIMARY};
                color: #FFFFFF;
                border: none;
                border-radius: {C.RADIUS_MD};
                padding: 8px 20px;
                font-size: 13px;
                font-weight: 600;
            }}
            QPushButton#primarySmall:hover {{
                background: {C.PRIMARY_HOVER};
            }}
            QTableWidget {{
                background: {C.BG_CARD};
                alternate-background-color: {C.FILL1};
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_MD};
                font-size: 13px;
                gridline-color: {C.BORDER_LIGHT};
                selection-background-color: {C.PRIMARY_LIGHT};
                selection-color: {C.PRIMARY};
            }}
            QTableWidget::item {{
                padding: 10px 14px;
                border-bottom: 1px solid {C.BORDER_LIGHT};
            }}
            QHeaderView::section {{
                background: {C.FILL1};
                color: {C.TEXT_REGULAR};
                padding: 10px 14px;
                border: none;
                border-bottom: 1px solid {C.BORDER};
                font-weight: 600;
                font-size: 13px;
            }}
            QComboBox {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER};
                border-radius: {C.RADIUS_MD};
                padding: 8px 36px 8px 14px;
                font-size: 13px;
                color: {C.TEXT_PRIMARY};
                min-height: 20px;
            }}
            QComboBox:hover {{
                border-color: {C.PRIMARY_HOVER};
            }}
            QComboBox QAbstractItemView {{
                background: {C.BG_ELEVATION};
                color: {C.TEXT_PRIMARY};
                border: 1px solid {C.BORDER};
                border-radius: {C.RADIUS_MD};
                selection-background-color: {C.PRIMARY_LIGHT};
                selection-color: {C.PRIMARY};
            }}
            QTextEdit {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER};
                border-radius: {C.RADIUS_MD};
                padding: 10px 14px;
                font-size: 13px;
                color: {C.TEXT_PRIMARY};
            }}
            QTextEdit:focus {{
                border-color: {C.PRIMARY};
                border-width: 2px;
                padding: 9px 13px;
            }}
        """)

        main_layout = QVBoxLayout(dialog)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 标题区
        header_layout = QHBoxLayout()
        header_layout.setSpacing(12)
        title_lbl = QLabel("术语库")
        title_lbl.setFont(QFont(C.FONT_FAMILY.split(",")[0].strip().strip("'"), 18, QFont.Weight.Bold))
        title_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; border: none; background: transparent;")
        header_layout.addWidget(title_lbl)

        data = load_glossary()
        stats = data.get("stats", {})
        stats_lbl = QLabel(f"{stats.get('total_terms', 0)} 条术语  ·  {stats.get('total_pairs', 0)} 个语言对")
        stats_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; border: none; background: transparent; border: 1px solid {C.BORDER}; border-radius: {C.RADIUS_SM}; padding: 3px 10px; background: {C.FILL1};")
        header_layout.addWidget(stats_lbl)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # 语言对选择行
        lang_card = QFrame()
        lang_card.setStyleSheet(f"""
            QFrame {{
                background: {C.BG_CARD};
                border: 1px solid {C.BORDER_LIGHT};
                border-radius: {C.RADIUS_MD};
                padding: 12px 16px;
            }}
        """)
        lang_row = QHBoxLayout(lang_card)
        lang_row.setContentsMargins(16, 10, 16, 10)
        lang_row.setSpacing(12)

        lang_lbl = QLabel("语言对")
        lang_lbl.setStyleSheet(f"color: {C.TEXT_REGULAR}; font-weight: 600; font-size: 13px; border: none; background: transparent;")
        lang_row.addWidget(lang_lbl)

        pairs = get_all_lang_pairs()
        pair_combo = QComboBox()
        pair_combo.setMinimumWidth(220)
        pair_combo.setFixedHeight(36)
        if pairs:
            for p in pairs:
                pair_combo.addItem(f"{p['key']}  ({p['count']} 条)", p['key'])
        else:
            pair_combo.addItem("暂无数据", "")
        lang_row.addWidget(pair_combo)
        lang_row.addStretch()

        import_btn = QPushButton("  导入术语  ")
        import_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        import_btn.setFixedHeight(34)
        lang_row.addWidget(import_btn)
        main_layout.addWidget(lang_card)

        # 表格
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["原文", "译文", "分类", "添加时间"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setEditTriggers(QTableWidget.SelectionTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        main_layout.addWidget(table)

        # 操作按钮行
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        def refresh_table():
            pair_key = pair_combo.currentData() or ""
            if not pair_key:
                table.setRowCount(0)
                return
            parts = pair_key.split("→")
            sl, tl = parts[0], parts[1] if len(parts) > 1 else ""
            terms = get_terms(sl, tl)
            table.setRowCount(len(terms))
            for row_idx, t in enumerate(terms):
                table.setItem(row_idx, 0, QTableWidgetItem(t.get("source", "")))
                table.setItem(row_idx, 1, QTableWidgetItem(t.get("target", "")))
                table.setItem(row_idx, 2, QTableWidgetItem(t.get("category", "")))
                table.setItem(row_idx, 3, QTableWidgetItem(t.get("added_at", "")))
            data = load_glossary()
            s = data.get("stats", {})
            stats_lbl.setText(f"{s.get('total_terms', 0)} 条术语  ·  {s.get('total_pairs', 0)} 个语言对")

        refresh_btn = QPushButton("  刷新  ")
        refresh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        refresh_btn.setFixedHeight(34)
        refresh_btn.clicked.connect(refresh_table)
        btn_row.addWidget(refresh_btn)

        delete_selected_btn = QPushButton("  删除选中  ")
        delete_selected_btn.setObjectName("dangerBtn")
        delete_selected_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        delete_selected_btn.setFixedHeight(34)

        def delete_selected():
            pair_key = pair_combo.currentData() or ""
            if not pair_key:
                return
            parts = pair_key.split("→")
            sl, tl = parts[0], parts[1] if len(parts) > 1 else ""
            rows = set(item.row() for item in table.selectedItems())
            for row in sorted(rows, reverse=True):
                src = table.item(row, 0).text() if table.item(row, 0) else ""
                delete_term(sl, tl, src)
            refresh_table()

        delete_selected_btn.clicked.connect(delete_selected)
        btn_row.addWidget(delete_selected_btn)

        clear_btn = QPushButton("  清空当前语言对  ")
        clear_btn.setObjectName("dangerBtn")
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.setFixedHeight(34)

        def clear_pair():
            pair_key = pair_combo.currentData() or ""
            if not pair_key:
                return
            ret = QMessageBox.question(dialog, "确认",
                f"确定清空 {pair_key} 的所有术语？此操作不可撤销。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if ret == QMessageBox.StandardButton.Yes:
                parts = pair_key.split("→")
                delete_lang_pair(parts[0], parts[1])
                refresh_table()

        clear_btn.clicked.connect(clear_pair)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        main_layout.addLayout(btn_row)

        # 导入术语对话框
        def do_import():
            imp_dialog = QDialog(dialog)
            imp_dialog.setWindowTitle("导入术语")
            imp_dialog.setMinimumSize(520, 380)
            imp_dialog.setStyleSheet(f"""
                QDialog {{
                    background-color: {C.BG_PAGE};
                    color: {C.TEXT_PRIMARY};
                }}
                QLabel {{
                    font-size: 13px;
                    color: {C.TEXT_PRIMARY};
                    border: none; background: transparent;
                }}
                QTextEdit {{
                    background: {C.BG_CARD};
                    border: 1px solid {C.BORDER};
                    border-radius: {C.RADIUS_MD};
                    padding: 10px 14px;
                    font-size: 13px;
                    color: {C.TEXT_PRIMARY};
                }}
                QTextEdit:focus {{
                    border-color: {C.PRIMARY};
                }}
                QPushButton#primarySmall {{
                    background: {C.PRIMARY};
                    color: #FFFFFF;
                    border: none;
                    border-radius: {C.RADIUS_MD};
                    padding: 8px 20px;
                    font-size: 13px;
                    font-weight: 600;
                }}
                QPushButton#primarySmall:hover {{
                    background: {C.PRIMARY_HOVER};
                }}
            """)

            imp_card = QFrame()
            imp_card.setStyleSheet(f"""
                QFrame {{
                    background: {C.BG_CARD};
                    border: 1px solid {C.BORDER_LIGHT};
                    border-radius: {C.RADIUS_LG};
                }}
            """)
            make_shadow_widget(imp_card, "sm")

            imp_outer = QVBoxLayout(imp_dialog)
            imp_outer.setContentsMargins(20, 20, 20, 20)
            imp_outer.addWidget(imp_card)

            imp_layout = QVBoxLayout(imp_card)
            imp_layout.setSpacing(12)
            imp_layout.setContentsMargins(24, 24, 24, 24)

            imp_title = QLabel("导入术语")
            imp_title.setFont(QFont(C.FONT_FAMILY.split(",")[0].strip().strip("'"), 15, QFont.Weight.Bold))
            imp_title.setStyleSheet(f"color: {C.TEXT_PRIMARY}; border: none; background: transparent;")
            imp_layout.addWidget(imp_title)

            imp_hint = QLabel("每行一条，格式：原文 → 译文")
            imp_hint.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; border: none; background: transparent;")
            imp_layout.addWidget(imp_hint)

            imp_text = QTextEdit()
            imp_text.setPlaceholderText("machine learning → 机器学习\nneural network → 神经网络\nthe state of the art → 最先进的")
            imp_text.setMinimumHeight(160)
            imp_layout.addWidget(imp_text)

            imp_result = QLabel("")
            imp_result.setStyleSheet(f"color: {C.SUCCESS}; font-size: 13px; border: none; background: transparent;")
            imp_layout.addWidget(imp_result)

            imp_btn_row = QHBoxLayout()
            imp_btn_row.addStretch()

            def do_import_text():
                pair_key = pair_combo.currentData() or ""
                if not pair_key:
                    return
                parts = pair_key.split("→")
                sl, tl = parts[0], parts[1] if len(parts) > 1 else ""
                added, total = import_glossary_from_text(
                    imp_text.toPlainText(), sl, tl)
                imp_result.setText(f"导入完成：新增 {added} 条 / 共解析 {total} 条")
                refresh_table()

            imp_btn = QPushButton("  确认导入  ")
            imp_btn.setObjectName("primarySmall")
            imp_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            imp_btn.clicked.connect(do_import_text)
            imp_btn_row.addWidget(imp_btn)
            imp_layout.addLayout(imp_btn_row)

            imp_dialog.exec()

        import_btn.clicked.connect(do_import)

        refresh_table()
        dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
