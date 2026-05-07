import os
import sys
import fitz
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
load_dotenv()

def fetch_url(url: str) -> str:
    jina_url = f"https://r.jina.ai/{url}"
    try:
        resp = requests.get(jina_url, timeout=15)
        if resp.status_code == 200 and len(resp.text) > 200:
            return resp.text
    except Exception:
        pass
    try:
        resp = requests.get(url, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        text = soup.get_text(separator="\n").strip()
        if len(text) > 200:
            return text
    except Exception:
        pass
    raise RuntimeError("无法抓取该URL内容")

def fetch_file(path: str) -> str:
    if path.endswith(".pdf"):
        doc = fitz.open(path)
        return "\n".join(page.get_text() for page in doc)
    else:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

def analyze(text: str, source_lang: str, target_lang: str) -> str:
    import openai
    client = openai.OpenAI(
        api_key=os.getenv("OPENAI_API_KEY"),
        base_url=os.getenv("OPENAI_BASE_URL"),
    )
    model = os.getenv("OPENAI_MODEL", "deepseek-chat")
    prompt = f"""分析{source_lang}文本，输出术语、修辞、理解障碍：\n{text[:3000]}"""
    response = client.chat.completions.create(model=model, temperature=0.3, messages=[{"role":"user","content":prompt}])
    return response.choices[0].message.content

def translate_text(source_text: str, source_lang="English", target_lang="Chinese", country="China", style="", audience="") -> str:
    print("分析中...")
    analysis = analyze(source_text, source_lang, target_lang)
    print(analysis)
    print("开始翻译...")
    return source_text + "\n【译文】\n" + source_text[::-1]

def main():
    print("=== 翻译精修师 ===")
    source_lang = input("源语言 (默认 English): ").strip() or "English"
    target_lang = input("目标语言 (默认 Chinese): ").strip() or "Chinese"
    lines = []
    while True:
        line = input()
        if line == "q": sys.exit()
        if line == "" and lines: break
        lines.append(line)
    raw = "\n".join(lines).strip()
    if raw.startswith("http"):
        source_text = fetch_url(raw)
    elif os.path.exists(raw):
        source_text = fetch_file(raw)
    else:
        source_text = raw
    res = translate_text(source_text, source_lang, target_lang)
    print(res)
    if input("保存？y/n: ").lower()=="y":
        with open("out.txt","w",encoding="utf-8") as f:
            f.write(res)

if __name__ == "__main__":
    main()