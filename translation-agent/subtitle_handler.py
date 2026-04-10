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
