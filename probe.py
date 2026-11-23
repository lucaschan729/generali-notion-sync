import io
import re
from urllib.parse import urljoin

import pdfplumber
import requests

PAGES = {
    "忠意人壽（香港）": "https://www.generali.com.hk/ZH_HK/claims_and_support/reference/generali_life",
    "忠意保險（香港分行）": "https://www.generali.com.hk/ZH_HK/claims_and_support/reference/generali_spa",
}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
KEYS = ("pdf", "fulfil", "ratio", "實現", "分紅", "現金價值", "cash", "download")
MAX_PDF = 12


def fetch(url):
    note = ""
    try:
        r = requests.get(url, headers=UA, timeout=60)
        if r.status_code == 200:
            return r.content, "普通連線"
        note = f"普通連線 {r.status_code}"
    except Exception as e:
        note = f"普通連線 {type(e).__name__}"
    try:
        from curl_cffi import requests as cr
        for imp in ("chrome", "safari"):
            r = cr.get(url, impersonate=imp, timeout=60)
            if r.status_code == 200:
                return r.content, f"curl_cffi {imp}"
            note += f"｜{imp} {r.status_code}"
    except Exception as e:
        note += f"｜curl_cffi {type(e).__name__}"
    return None, note


def links(html, base):
    out, seen = [], set()
    for href, inner in re.findall(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html, flags=re.S | re.I):
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", inner)).strip()
        url = urljoin(base, href.strip())
        blob = (url + " " + text).lower()
        if url in seen or not any(k in blob for k in KEYS):
            continue
        seen.add(url)
        out.append((url, text))
    return out


def show_pdf(url, text):
    data, how = fetch(url)
    if data is None:
        print(f"   ✗ 下載失敗（{how}）")
        return
    if not data.startswith(b"%PDF"):
        print(f"   － 不是 PDF（{len(data) // 1024} KB，{how}）")
        return
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        meta = (pdf.metadata or {}).get("CreationDate", "")
        texts = [p.extract_text() or "" for p in pdf.pages]
        full = "\n".join(texts)
        print(f"   ✓ PDF {len(data) // 1024} KB｜{len(pdf.pages)} 頁｜建立 {meta}｜{how}")
        print(f"   含「總現金價值」：{'總現金價值' in full or 'total cash value' in full.lower()}"
              f"｜含「實現率」：{'實現率' in full or 'fulfil' in full.lower()}")
        print("   --- 文字（前 60 行） ---")
        for line in full.splitlines()[:60]:
            print("   " + line)
        for i, page in enumerate(pdf.pages[:3], 1):
            for j, table in enumerate(page.extract_tables(), 1):
                cols = max(len(r) for r in table) if table else 0
                print(f"   --- 第 {i} 頁 表格 {j}：{len(table)} 行 × {cols} 欄（前 8 行） ---")
                for row in table[:8]:
                    print("   " + " | ".join(str(c or "").replace("\n", " ") for c in row))


def main():
    for name, url in PAGES.items():
        print(f"=== {name} ===")
        print(url)
        data, how = fetch(url)
        if data is None:
            print(f"✗ 頁面讀取失敗（{how}）")
            print()
            continue
        html = data.decode("utf-8", errors="replace")
        print(f"✓ 頁面 {len(html) // 1024} KB（{how}）｜HTML 表格 {len(re.findall(r'<table', html, flags=re.I))} 個")
        found = links(html, url)
        print(f"相關連結 {len(found)} 條：")
        for u, t in found:
            print(f" - {t[:60]} → {u}")
        print()
        count = 0
        for u, t in found:
            if count >= MAX_PDF:
                print(f"（只試前 {MAX_PDF} 條，其餘略過）")
                break
            count += 1
            print(f">> {t[:60]}")
            print(f"   {u}")
            try:
                show_pdf(u, t)
            except Exception as e:
                print(f"   ✗ 讀取出錯：{type(e).__name__} {e}")
            print()
    print("探查完成（只讀，沒有上傳）")


if __name__ == "__main__":
    main()
