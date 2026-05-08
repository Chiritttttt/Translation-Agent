"""
Translation Agent — 命令行入口
================================
五步翻译工作流：分析 → 构建提示 → 初译 → 审校 → 终稿润色
依赖 translation-agent/gui.py 中的核心函数。
"""

import os
import sys

# 将脚本所在目录加入搜索路径，以便导入同目录模块
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from dotenv import load_dotenv
load_dotenv()

import openai

_api_key = os.getenv("OPENAI_API_KEY", "")
_base_url = os.getenv("OPENAI_BASE_URL", "") or None
_model_name = os.getenv("OPENAI_MODEL", "deepseek-chat")
_max_tokens = int(os.getenv("OPENAI_MAX_TOKENS", "16384"))

if not _api_key:
    print("错误：请在 .env 文件中设置 OPENAI_API_KEY")
    sys.exit(1)

client = openai.OpenAI(
    api_key=_api_key,
    base_url=_base_url,
)


def chat(system, user, temperature=0.3):
    """调用 LLM"""
    resp = client.chat.completions.create(
        model=_model_name,
        temperature=temperature,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        max_tokens=_max_tokens,
    )
    return resp.choices[0].message.content


# ═══ 导入 gui.py 中的五步翻译函数（复用核心逻辑） ═══
from gui import (
    step1_analyze,
    step2_build_prompt,
    step3_draft,
    step4_critique,
    step5_final,
    fetch_url,
)
from file_handler import read_file


def translate_file_or_text(source_text, source_lang="English",
                           target_lang="Chinese", style="",
                           audience=""):
    """执行完整的五步翻译流程"""
    print("第一步：深度分析源文本...")
    analysis = step1_analyze(source_text, source_lang, target_lang,
                              audience, style)

    print("\n── 分析报告 ──")
    print(analysis)
    print()

    print("第二步：组装翻译提示...")
    prompt = step2_build_prompt(analysis, source_lang, target_lang,
                                 audience, style)

    print("第三步：初译...")
    draft = step3_draft(source_text, prompt, source_lang, target_lang)

    print("第四步：审校...")
    critique = step4_critique(source_text, draft, analysis,
                               source_lang, target_lang)

    print("\n── 审校报告 ──")
    print(critique)
    print()

    print("第五步：终稿润色...")
    final = step5_final(source_text, draft, critique, target_lang)

    return final, analysis, critique


def main():
    print("=== Translation Agent (CLI) ===")
    print()

    source_lang = input("源语言 (默认 English): ").strip() or "English"
    target_lang = input("目标语言 (默认 Chinese): ").strip() or "Chinese"
    style = input("翻译风格 (默认 auto): ").strip() or "auto"
    audience = input("目标读者 (默认 一般读者): ").strip() or "一般读者"
    print()

    # 读取输入
    print("请输入要翻译的文本（输入空行结束，或输入 URL / 文件路径）：")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.lower() == "q":
            sys.exit(0)
        if line == "" and lines:
            break
        lines.append(line)

    raw = "\n".join(lines).strip()
    if not raw:
        print("未输入内容，退出。")
        sys.exit(0)

    # 判断输入类型
    if raw.startswith("http://") or raw.startswith("https://"):
        print(f"正在抓取 URL: {raw}")
        source_text = fetch_url(raw)
    elif os.path.exists(raw):
        print(f"正在读取文件: {raw}")
        file_type, source_text, _ = read_file(raw)
        print(f"文件类型: {file_type}, 共 {len(source_text)} 字符")
    else:
        source_text = raw

    print(f"\n原文共 {len(source_text)} 字符\n")

    # 执行翻译
    try:
        result, analysis, critique = translate_file_or_text(
            source_text, source_lang, target_lang, style, audience
        )
    except Exception as e:
        print(f"\n翻译失败: {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("最终译文：")
    print("=" * 60)
    print(result)

    # 保存选项
    save = input("\n保存译文？(y/n, 默认 y): ").strip().lower()
    if save != "n":
        if raw.startswith("http"):
            out_name = "translated_url.txt"
        elif os.path.exists(raw):
            base = os.path.splitext(os.path.basename(raw))[0]
            out_name = f"{base}_{target_lang}.txt"
        else:
            out_name = "translated.txt"

        with open(out_name, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"已保存到: {out_name}")

        # 同时保存分析报告
        report_name = out_name.replace(".txt", "_report.txt")
        with open(report_name, "w", encoding="utf-8") as f:
            f.write("## 分析报告\n\n")
            f.write(analysis)
            f.write("\n\n## 审校报告\n\n")
            f.write(critique)
        print(f"分析/审校报告已保存到: {report_name}")


if __name__ == "__main__":
    main()
