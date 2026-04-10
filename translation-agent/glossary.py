"""
术语库管理模块
==============
- 本地 JSON 文件存储，路径: ~/.translation_agent/glossary.json
- 支持按语言对（source→target）分组
- 自动去重
- 从 AI 分析报告中自动提取术语
"""

import os
import re
import json
from datetime import datetime
from pathlib import Path

GLOSSARY_DIR = Path.home() / ".translation_agent"
GLOSSARY_FILE = GLOSSARY_DIR / "glossary.json"


# ─── 数据结构 ────────────────────────────────────────────

def _ensure_dir():
    """确保术语库目录存在"""
    GLOSSARY_DIR.mkdir(parents=True, exist_ok=True)


def load_glossary():
    """加载完整术语库，返回 dict 结构:
    {
        "lang_pairs": {
            "en→zh": {
                "terms": [
                    {"source": "machine learning", "target": "机器学习", "category": "术语", "added_at": "..."},
                    ...
                ]
            }
        },
        "stats": {"total_terms": 42, "total_pairs": 5}
    }
    """
    _ensure_dir()
    if not GLOSSARY_FILE.exists():
        empty = {"lang_pairs": {}, "stats": {"total_terms": 0, "total_pairs": 0}}
        save_glossary(empty)
        return empty

    try:
        with open(GLOSSARY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 兼容旧格式
        if "lang_pairs" not in data:
            data = {"lang_pairs": {}, "stats": {"total_terms": 0, "total_pairs": 0}}
        return data
    except (json.JSONDecodeError, KeyError):
        empty = {"lang_pairs": {}, "stats": {"total_terms": 0, "total_pairs": 0}}
        save_glossary(empty)
        return empty


def save_glossary(data):
    """保存术语库到磁盘"""
    _ensure_dir()
    # 更新统计
    total_terms = 0
    total_pairs = 0
    for pair_key, pair_data in data.get("lang_pairs", {}).items():
        terms = pair_data.get("terms", [])
        total_terms += len(terms)
        if terms:
            total_pairs += 1
    data["stats"] = {"total_terms": total_terms, "total_pairs": total_pairs}

    with open(GLOSSARY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ─── 语言对 key ──────────────────────────────────────────

def _pair_key(source_lang, target_lang):
    """生成语言对 key，如 'en→zh'"""
    lang_map = {
        "english": "en", "chinese": "zh", "japanese": "ja", "korean": "ko",
        "french": "fr", "german": "de", "spanish": "es",
    }
    sl = lang_map.get(source_lang.lower(), source_lang.lower()[:2])
    tl = lang_map.get(target_lang.lower(), target_lang.lower()[:2])
    return f"{sl}→{tl}"


# ─── 添加术语 ────────────────────────────────────────────

def add_term(source, target, source_lang, target_lang, category="术语"):
    """添加单条术语到术语库，自动去重。

    Args:
        source: 原文词汇
        target: 译文
        source_lang: 源语言
        target_lang: 目标语言
        category: 分类（术语 / 表达 / 缩写 等）

    Returns:
        bool: 是否新增（False 表示已存在）
    """
    data = load_glossary()
    key = _pair_key(source_lang, target_lang)

    if key not in data["lang_pairs"]:
        data["lang_pairs"][key] = {"terms": [], "source_lang": source_lang, "target_lang": target_lang}

    # 去重：同语言对 + 同原文 + 同译文 视为重复
    source_clean = source.strip().lower()
    for existing in data["lang_pairs"][key]["terms"]:
        if existing["source"].strip().lower() == source_clean:
            return False  # 已存在

    data["lang_pairs"][key]["terms"].append({
        "source": source.strip(),
        "target": target.strip(),
        "category": category,
        "added_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    })

    save_glossary(data)
    return True


def add_terms_batch(terms, source_lang, target_lang):
    """批量添加术语。

    Args:
        terms: list of dict, 每个包含 {"source": str, "target": str, "category": str}
        source_lang / target_lang: 语言

    Returns:
        (added_count, skipped_count)
    """
    data = load_glossary()
    key = _pair_key(source_lang, target_lang)

    if key not in data["lang_pairs"]:
        data["lang_pairs"][key] = {"terms": [], "source_lang": source_lang, "target_lang": target_lang}

    existing_sources = {
        t["source"].strip().lower() for t in data["lang_pairs"][key]["terms"]
    }

    added = 0
    skipped = 0
    for item in terms:
        src = item.get("source", "").strip()
        tgt = item.get("target", "").strip()
        cat = item.get("category", "术语")

        if not src or not tgt:
            skipped += 1
            continue

        if src.lower() in existing_sources:
            skipped += 1
            continue

        data["lang_pairs"][key]["terms"].append({
            "source": src,
            "target": tgt,
            "category": cat,
            "added_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
        existing_sources.add(src.lower())
        added += 1

    save_glossary(data)
    return added, skipped


# ─── 从 AI 分析报告中提取术语 ────────────────────────────

def extract_terms_from_analysis(analysis_text, source_lang, target_lang):
    """从 step1_analyze 输出的分析报告中提取术语。

    AI 输出格式示例：
    ## 术语表（原文 → 译法）
    - machine learning → 机器学习
    - neural network → 神经网络
    或：
    - machine learning / 机器学习
    或：
    **machine learning**: 机器学习

    Returns:
        list of dict: [{"source": ..., "target": ..., "category": ...}, ...]
    """
    terms = []

    # 多种匹配模式
    patterns = [
        # "- source → target" 或 "- source → target"
        re.compile(r"^[-•]\s*(.+?)\s*[→➡>]\s*(.+)$", re.MULTILINE),
        # "- source / target" 或 "- source // target"
        re.compile(r"^[-•]\s*(.+?)\s*[/／]\s*(.+)$", re.MULTILINE),
        # "**source**: target" 或 "source: target"（行首，排除 ## 等标题）
        re.compile(r"^\s{0,6}(\S.{1,60}?)\s*[:：]\s*(.{1,60})$", re.MULTILINE),
    ]

    def _extract_from_lines(lines_to_process):
        """从行列表中提取术语（复用逻辑，避免重复代码）"""
        result = []
        for line in lines_to_process:
            for pattern in patterns:
                match = pattern.match(line)
                if match:
                    src = match.group(1).strip().strip("*").strip()
                    tgt = match.group(2).strip().strip("*").strip()
                    if len(src) <= 80 and len(tgt) <= 80 and not src.startswith("#"):
                        src = re.sub(r"^[`*\-]+\s*", "", src)
                        src = re.sub(r"\s*[`*]+$", "", src)
                        tgt = re.sub(r"^[`*\-]+\s*", "", tgt)
                        tgt = re.sub(r"\s*[`*]+$", "", tgt)
                        if src and tgt and src.lower() != tgt.lower():
                            cat = "术语"
                            if any(w in src.lower() for w in ["the ", "a ", "an ", "is ", "are ", " to "]):
                                cat = "表达"
                            result.append({"source": src, "target": tgt, "category": cat})
                    break
        return result

    # ── 第一阶段：从术语表段落中提取 ──
    lines = analysis_text.split("\n")
    in_terms_section = False
    term_lines = []

    for line in lines:
        stripped = line.strip()
        # Bug fix: 放宽术语段落检测，兼容分段分析的 "### 第 N/M 段补充术语" 格式
        if "术语" in stripped or ("词汇" in stripped) or ("表达" in stripped and "惯用" in stripped):
            # 排除描述性文字（如"这些术语很重要"），必须是标题行
            if stripped.startswith("#") or "表" in stripped or "补充" in stripped:
                in_terms_section = True
                continue
        # 检测下一个 ## / ### 非术语标题，表示段落结束
        if in_terms_section and (stripped.startswith("## ") or stripped.startswith("### ")):
            if "术语" in stripped or "词汇" in stripped:
                continue  # 仍是术语相关标题（如 "### 第 2/3 段补充术语"）
            else:
                in_terms_section = False
                continue
        if in_terms_section and stripped:
            if any(c in stripped for c in ["→", "➡", ">", "：", ":", "／", "/"]) and not stripped.startswith("#"):
                term_lines.append(stripped)

    terms = _extract_from_lines(term_lines)

    # ── 第二阶段：全文兜底匹配 ──
    # 段落检测可能遗漏非标准格式，对全文做一次兜底
    all_lines = [l.strip() for l in analysis_text.split("\n")]
    fallback_lines = [l for l in all_lines
                      if l.startswith("-") or l.startswith("•")
                      if any(c in l for c in ["→", "➡", ">", "/", "／"])
                      if len(l) <= 200]
    terms.extend(_extract_from_lines(fallback_lines))

    # 去重
    seen = set()
    unique_terms = []
    for t in terms:
        key = f"{t['source'].lower()}|{t['target'].lower()}"
        if key not in seen:
            seen.add(key)
            unique_terms.append(t)

    return unique_terms


# ─── 查询术语库 ──────────────────────────────────────────

def get_terms(source_lang, target_lang, category=None):
    """获取指定语言对的术语列表。

    Args:
        category: 可选，按分类过滤（术语 / 表达 / 缩写 等）

    Returns:
        list of dict
    """
    data = load_glossary()
    key = _pair_key(source_lang, target_lang)
    pair = data["lang_pairs"].get(key, {})
    terms = pair.get("terms", [])

    if category:
        terms = [t for t in terms if t.get("category") == category]

    return terms


def get_glossary_prompt_block(source_lang, target_lang):
    """生成翻译提示词中的术语约束段落。

    Returns:
        str: 如有术语则返回提示词片段，否则返回空字符串
    """
    terms = get_terms(source_lang, target_lang)
    if not terms:
        return ""

    lines = ["## Glossary (术语约束 — 必须严格遵守)"]
    lines.append("The following terms MUST be translated exactly as specified:")
    lines.append("")

    for t in terms:
        cat = t.get("category", "术语")
        lines.append(f"- {t['source']} → {t['target']}  [{cat}]")

    lines.append("")
    lines.append("IMPORTANT: Never deviate from the above glossary. These are established terminology standards.")

    return "\n".join(lines)


def get_all_lang_pairs():
    """获取所有语言对及其术语数量。

    Returns:
        list of dict: [{"key": "en→zh", "source_lang": ..., "target_lang": ..., "count": N}, ...]
    """
    data = load_glossary()
    result = []
    for key, pair in data.get("lang_pairs", {}).items():
        result.append({
            "key": key,
            "source_lang": pair.get("source_lang", ""),
            "target_lang": pair.get("target_lang", ""),
            "count": len(pair.get("terms", [])),
        })
    return sorted(result, key=lambda x: x["count"], reverse=True)


# ─── 删除 / 编辑 ─────────────────────────────────────────

def edit_term(source_lang, target_lang, source_text, new_target=None, new_source=None, new_category=None):
    """编辑指定术语。

    Args:
        source_lang / target_lang: 语言对
        source_text: 原术语的原文（用于定位）
        new_target: 新译文（None 则不修改）
        new_source: 新原文（None 则不修改）
        new_category: 新分类（None 则不修改）

    Returns:
        bool: 是否修改成功
    """
    data = load_glossary()
    key = _pair_key(source_lang, target_lang)

    if key not in data["lang_pairs"]:
        return False

    source_clean = source_text.strip().lower()
    for t in data["lang_pairs"][key]["terms"]:
        if t["source"].strip().lower() == source_clean:
            if new_target is not None and new_target.strip():
                t["target"] = new_target.strip()
            if new_source is not None and new_source.strip():
                t["source"] = new_source.strip()
            if new_category is not None and new_category.strip():
                t["category"] = new_category.strip()
            save_glossary(data)
            return True
    return False


def delete_term(source_lang, target_lang, source_text):
    """删除指定术语。

    Returns:
        bool: 是否删除成功
    """
    data = load_glossary()
    key = _pair_key(source_lang, target_lang)

    if key not in data["lang_pairs"]:
        return False

    original_count = len(data["lang_pairs"][key]["terms"])
    data["lang_pairs"][key]["terms"] = [
        t for t in data["lang_pairs"][key]["terms"]
        if t["source"].strip().lower() != source_text.strip().lower()
    ]

    if len(data["lang_pairs"][key]["terms"]) < original_count:
        save_glossary(data)
        return True
    return False


def delete_lang_pair(source_lang, target_lang):
    """删除整个语言对的所有术语。"""
    data = load_glossary()
    key = _pair_key(source_lang, target_lang)
    if key in data["lang_pairs"]:
        del data["lang_pairs"][key]
        save_glossary(data)


def clear_all():
    """清空整个术语库。"""
    empty = {"lang_pairs": {}, "stats": {"total_terms": 0, "total_pairs": 0}}
    save_glossary(empty)


# ─── 导出 ────────────────────────────────────────────────

def export_glossary_text(source_lang, target_lang, category=None):
    """导出术语为纯文本格式。

    Returns:
        str
    """
    terms = get_terms(source_lang, target_lang, category)
    if not terms:
        return "（术语库为空）"

    lines = [f"术语表 ({source_lang} → {target_lang})"]
    lines.append("=" * 50)
    lines.append("")

    # 按分类分组
    categories = {}
    for t in terms:
        cat = t.get("category", "术语")
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(t)

    for cat, cat_terms in categories.items():
        lines.append(f"【{cat}】")
        for t in cat_terms:
            lines.append(f"  {t['source']}  →  {t['target']}")
        lines.append("")

    return "\n".join(lines)


def export_glossary_json(source_lang=None, target_lang=None):
    """导出术语库为 JSON 格式。

    Args:
        source_lang/target_lang: 如果指定则只导出该语言对，否则导出全部

    Returns:
        str (JSON)
    """
    data = load_glossary()

    if source_lang and target_lang:
        key = _pair_key(source_lang, target_lang)
        pair = data["lang_pairs"].get(key, {"terms": []})
        return json.dumps(pair.get("terms", []), ensure_ascii=False, indent=2)
    else:
        return json.dumps(data, ensure_ascii=False, indent=2)


def import_glossary_from_text(text, source_lang, target_lang):
    """从文本导入术语，每行格式: "原文 → 译文" 或 "原文 / 译文"。

    Returns:
        (added_count, total_count)
    """
    terms = []
    for line in text.strip().split("\n"):
        line = line.strip()
        if not line or line.startswith("#") or line.startswith("="):
            continue

        # 尝试多种分隔符
        for sep in ["→", "➡", ">", "：", ":", "／", "/"]:
            if sep in line:
                parts = line.split(sep, 1)
                if len(parts) == 2:
                    src = parts[0].strip().strip("-•*").strip()
                    tgt = parts[1].strip().strip("-•*").strip()
                    if src and tgt and len(src) <= 80 and len(tgt) <= 80:
                        cat = "表达" if any(w in src.lower() for w in ["the ", " a ", " to "]) else "术语"
                        terms.append({"source": src, "target": tgt, "category": cat})
                    break

    if not terms:
        return 0, 0

    added, skipped = add_terms_batch(terms, source_lang, target_lang)
    return added, len(terms)


def clear_glossary(source_lang=None, target_lang=None):
    """清空术语库（复用统一的 GLOSSARY_FILE 路径）"""
    if source_lang and target_lang:
        data = load_glossary()
        key = _pair_key(source_lang, target_lang)
        if key in data.get("lang_pairs", {}):
            del data["lang_pairs"][key]
            save_glossary(data)
    else:
        # 清空全部
        clear_all()
