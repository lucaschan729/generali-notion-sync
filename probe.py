import collections
import io
import re
import statistics
from urllib.parse import urljoin

import pdfplumber
from curl_cffi import requests as cr

PAGES = [
    ("忠意人壽（香港）", "https://www.generali.com.hk/ZH_HK/claims_and_support/reference/generali_life"),
    ("忠意保險（香港分行）", "https://www.generali.com.hk/ZH_HK/claims_and_support/reference/generali_spa"),
]
LABELS = [str(i) for i in range(1, 11)] + ["10+"]
YEAR_RE = re.compile(r"^第(\d{1,2}\+?)個$")
PCT_RE = re.compile(r"^\d+(\.\d+)?%$")
NUM_RE = re.compile(r"^\d+(\.\d+)?$")
TAILS = {"停止銷售": "sales", "尚未推出": "launched", "不適用": "N/A"}
SHORT = {"停止銷售": "停售", "尚未推出": "未推出", "不適用": "不適用"}
FILLERS = {"Closed", "to", "sales", "Not", "yet", "launched", "N/A"}
SIMPLIFIED = "个实现产货红单险"


def fetch(url):
    notes = []
    for imp in ("chrome", "safari"):
        try:
            r = cr.get(url, impersonate=imp, timeout=60)
            if r.status_code == 200:
                return r.content
            notes.append(f"{imp} {r.status_code}")
        except Exception as e:
            notes.append(f"{imp} {type(e).__name__}")
    raise SystemExit(f"✗ 讀取失敗：{url}（{'｜'.join(notes)}）")


def find_pdf(html, base):
    for href, inner in re.findall(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html, flags=re.S | re.I):
        text = re.sub(r"\s+", "", re.sub(r"<[^>]+>", "", inner))
        if "實現率" in text and ".pdf" in href.lower():
            return urljoin(base, href.strip())
    return None


def has_cjk(s):
    return bool(re.search(r"[\u4e00-\u9fff]", s))


def is_en(s):
    return bool(re.search(r"[A-Za-z]", s)) and not has_cjk(s)


def cx(w):
    return (w["x0"] + w["x1"]) / 2


def clean_words(page):
    """讀出字詞；如果數字和 % 被拆開，合併回一個。"""
    out = []
    for w in sorted(page.extract_words(), key=lambda v: (round(v["top"]), v["x0"])):
        if (out and w["text"] == "%" and NUM_RE.match(out[-1]["text"])
                and abs(w["top"] - out[-1]["top"]) < 3 and 0 <= w["x0"] - out[-1]["x1"] < 6):
            out[-1] = dict(out[-1], text=out[-1]["text"] + "%", x1=w["x1"])
        else:
            out.append(w)
    return out


def lines_of(words):
    """把字詞按高度分成一行行，每行由左至右。"""
    out = []
    for w in sorted(words, key=lambda v: (v["top"], v["x0"])):
        if out and w["top"] - out[-1][0]["top"] <= 3:
            out[-1].append(w)
        else:
            out.append([w])
    return [sorted(line, key=lambda v: v["x0"]) for line in out]


def line_text(line):
    text = ""
    for w in line:
        if text and not (has_cjk(text[-1]) or has_cjk(w["text"][0])):
            text += " "
        text += w["text"]
    return text


def blocks_of(words, h):
    """把一欄的字詞分成一格格：中文行之後接英文行；遇到新的中文行或較大空隙當作新一格。"""
    out = []
    prev = None
    for line in lines_of(words):
        text = line_text(line)
        top = min(w["top"] for w in line)
        bottom = max(w["bottom"] for w in line)
        if prev is None or (has_cjk(text) and prev["en"]) or top - prev["bottom"] > 0.8 * h:
            out.append({"zh": [], "en": [], "top": top, "bottom": bottom})
        block = out[-1]
        (block["en"] if is_en(text) else block["zh"]).append(text)
        block["bottom"] = bottom
        prev = {"en": is_en(text), "bottom": bottom}
    for block in out:
        block["name"] = "".join(block["zh"]) or " ".join(block["en"])
        block["mid"] = (block["top"] + block["bottom"]) / 2
    return out


def bounds_of(centres):
    """由每行的中心位置推算行與行之間的分界。"""
    n = len(centres)
    if n == 1:
        return [centres[0] - 18, centres[0] + 18]
    mids = [(centres[i] + centres[i + 1]) / 2 for i in range(n - 1)]
    return [2 * centres[0] - mids[0]] + mids + [2 * centres[-1] - mids[-1]]


def assign(blocks, bounds):
    """把一欄的格（可能是合併格）配對到連續的行。回傳每行對應的格和最大偏差。"""
    n = len(bounds) - 1
    m = len(blocks)
    if m == 0 or m > n:
        return None, 0.0
    owner, start, worst = [], 0, 0.0
    for j, block in enumerate(blocks):
        if j == m - 1:
            end = n
        else:
            last = n - (m - 1 - j)
            target = 2 * block["mid"] - bounds[start]
            end = min(range(start + 1, last + 1), key=lambda e: abs(bounds[e] - target))
        worst = max(worst, abs((bounds[start] + bounds[end]) / 2 - block["mid"]))
        owner += [block] * (end - start)
        start = end
    return owner, worst


def dump(words, y0, y1, x_max, limit):
    for line in lines_of([w for w in words if y0 <= w["top"] < y1 and w["x0"] < x_max])[:limit]:
        print(f"      y={line[0]['top']:.0f} " + "  ".join(f"{w['x0']:.0f}:{w['text']}" for w in line))


def head_info(words, band):
    band_top = min(w["top"] for w in band)
    cols = {YEAR_RE.match(w["text"]).group(1): cx(w) for w in band}

    def above(name):
        c = [w for w in words if w["text"] == name and band_top - 60 < w["top"] <= band_top + 5]
        return max(c, key=lambda v: v["top"]) if c else None

    found = [above("產品系列"), above("紅利類別"), above("貨幣")]
    tops = [w["top"] for w in found if w]
    head_top = min(tops) if tops else band_top
    x_lim = min(cols.values())
    near = [w for w in words if head_top - 45 < w["top"] < head_top - 3 and w["x0"] < x_lim]
    cat_lines = [ln for ln in lines_of(near) if has_cjk(line_text(ln))]
    cat = line_text(cat_lines[-1]) if cat_lines else ""
    start = cat_lines[-1][0]["top"] if cat_lines else head_top
    return {"top": band_top, "cols": cols, "type": found[1], "cur": found[2], "cat": cat, "start": start}


def parse_band(page, words, info, end, pno, cat, sample):
    """讀一個表頭下面的資料。回傳（行、需要留意的數目）。"""
    cols = info["cols"]

    def fail(msg, y0, y1):
        print(f"   ✗ 第 {pno} 頁：{msg}")
        dump(words, y0, y1, page.width, 90)
        return [], 1

    if sorted(cols) != sorted(LABELS):
        return fail(f"表頭的保單年度欄不是 11 欄：{sorted(cols)}", info["top"] - 40, info["top"] + 40)
    if not info["type"] or not info["cur"]:
        return fail("找不到「紅利類別」或「貨幣」表頭", info["top"] - 60, info["top"] + 40)
    xs = [cols[k] for k in LABELS]
    pitch = statistics.median([xs[i + 1] - xs[i] for i in range(10)])
    x_left = xs[0] - pitch / 2
    b1 = info["type"]["x0"] - 4
    b2 = (info["type"]["x1"] + info["cur"]["x0"]) / 2
    notes = [w["top"] for w in words if w["text"].startswith("註") and info["top"] < w["top"] < end]
    if notes:
        end = min(notes)
    toks = [w for w in words if info["top"] < w["top"] < end and cx(w) >= x_left
            and (PCT_RE.match(w["text"]) or w["text"] in TAILS)]
    if not toks:
        return fail("表頭下面沒有讀到數值", info["top"] - 40, min(end, info["top"] + 300))
    h = statistics.median([w["bottom"] - w["top"] for w in toks])
    first = min(w["top"] for w in toks)
    closers = [w["bottom"] for w in words
               if re.match(r"^\d{4}\)$", w["text"]) and info["top"] < w["top"] < first]
    data_top = max(closers) if closers else first - 1.5 * h
    zone = [w for w in words if data_top - 1 <= w["top"] < end]

    percol = {k: [] for k in LABELS}
    for w in toks:
        k = min(LABELS, key=lambda v: abs(cols[v] - cx(w)))
        percol[k].append(w)
    counts = [len(percol[k]) for k in LABELS]
    if len(set(counts)) != 1:
        return fail(f"各保單年度欄讀到的格數不一致：{counts}", data_top - 5, end)
    n = counts[0]
    for k in LABELS:
        percol[k].sort(key=lambda v: v["top"])
    grid = [[percol[k][i] for k in LABELS] for i in range(n)]
    for i in range(n - 1):
        if max(w["top"] for w in grid[i]) >= min(w["top"] for w in grid[i + 1]):
            return fail(f"第 {i + 1} 行和第 {i + 2} 行的上下位置重疊", data_top - 5, end)

    def stack_bottom(w, k):
        tail = TAILS.get(w["text"])
        if not tail:
            return w["bottom"]
        c = [v for v in zone if v["text"] == tail and abs(cx(v) - cols[k]) <= pitch / 2
             and 0 < v["top"] - w["top"] < 4.5 * h]
        if c:
            return min(c, key=lambda v: v["top"])["bottom"]
        return w["bottom"] + (1.2 if tail == "N/A" else 2.4) * h

    bottoms = [[stack_bottom(w, k) for w, k in zip(row, LABELS)] for row in grid]
    centres = [statistics.median([(w["top"] + b) / 2 for w, b in zip(row, bs)])
               for row, bs in zip(grid, bottoms)]
    last_bottom = max(max(bs) for bs in bottoms)
    bad = 0
    odd = [w["text"] for w in zone if cx(w) >= x_left and w["top"] <= last_bottom
           and not (PCT_RE.match(w["text"]) or w["text"] in TAILS or w["text"] in FILLERS)]
    if odd:
        bad += 1
        print(f"   ⚠ 第 {pno} 頁：年度欄內有未能識別的文字：{odd[:12]}")

    left = [w for w in zone if cx(w) < x_left]
    bounds = bounds_of(centres)
    owners = []
    for label, lo, hi in (("產品系列", -1, b1), ("紅利類別", b1, b2), ("貨幣", b2, x_left)):
        blocks = [b for b in blocks_of([w for w in left if lo <= cx(w) < hi], h) if b["top"] < last_bottom]
        owner, worst = assign(blocks, bounds)
        if owner is None:
            return fail(f"「{label}」一欄讀到 {len(blocks)} 格，但數值有 {n} 行", data_top - 5, end)
        if worst > 0.8 * h:
            bad += 1
            print(f"   ⚠ 第 {pno} 頁：「{label}」與行的位置配對偏差 {worst:.0f}pt，可能配錯")
        owners.append(owner)

    n_prod = len({id(b) for b in owners[0]})
    print(f"   第 {pno} 頁｜{cat or '（未讀到種類）'}｜{n} 行｜產品 {n_prod} 格")
    rows = []
    for i in range(n):
        vals = [round(float(w["text"][:-1]) / 100, 4) if PCT_RE.match(w["text"]) else w["text"]
                for w in grid[i]]
        rows.append({"cat": cat, "prod": owners[0][i]["name"], "type": owners[1][i]["name"],
                     "cur": owners[2][i]["name"], "vals": vals, "page": pno})
    if sample or bad:
        print("   --- 左邊三欄的位置（供核對） ---")
        dump(left, data_top - 1, last_bottom + 30, x_left, 45)
    return rows, bad


def parse_pdf(data):
    rows, bad = [], 0
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        meta = (pdf.metadata or {}).get("CreationDate", "") or ""
        texts = [p.extract_text() or "" for p in pdf.pages]
        m = re.search(r"(\d{4})\s*報告年度", "\n".join(texts))
        year = m.group(1) if m else "？"
        print(f"   PDF {len(data) // 1024} KB｜{len(pdf.pages)} 頁｜報告年度 {year}｜建立 {meta[2:10]}")
        cat, sample, skipped = "", True, []
        for pno, page in enumerate(pdf.pages, 1):
            if sum(texts[pno - 1].count(ch) for ch in SIMPLIFIED) >= 3:
                skipped.append(pno)
                continue
            words = clean_words(page)
            heads = sorted([w for w in words if YEAR_RE.match(w["text"])], key=lambda v: v["top"])
            bands = []
            for w in heads:
                if bands and w["top"] - bands[-1][0]["top"] <= 4:
                    bands[-1].append(w)
                else:
                    bands.append([w])
            if not bands:
                if any(PCT_RE.match(w["text"]) or w["text"] in TAILS for w in words):
                    bad += 1
                    print(f"   ✗ 第 {pno} 頁：有數值但沒有表頭")
                    dump(words, 0, page.height, page.width, 60)
                continue
            infos = [head_info(words, b) for b in bands]
            for i, info in enumerate(infos):
                end = infos[i + 1]["start"] - 2 if i + 1 < len(infos) else page.height
                cat = info["cat"] or cat
                got, b = parse_band(page, words, info, end, pno, cat, sample)
                rows += got
                bad += b
                sample = False
        print(f"   略過簡體頁：{skipped}")
    return rows, bad


def show(v):
    return f"{v * 100:g}%" if isinstance(v, float) else SHORT[v]


def report(rows):
    for r in rows:
        print(f"   p{r['page']}｜{r['prod']}｜{r['type']}｜{r['cur']}｜" + " ".join(show(v) for v in r["vals"]))
    nums = [v for r in rows for v in r["vals"] if isinstance(v, float)]
    marks = collections.Counter(v for r in rows for v in r["vals"] if not isinstance(v, float))
    n_prod = len({(r["cat"], r["prod"]) for r in rows})
    print(f"   小結：產品 {n_prod} 個｜{len(rows)} 行｜數值 {len(nums)} 個")
    if nums:
        low = sum(1 for v in nums if v < 0.8)
        high = sum(1 for v in nums if v >= 1)
        print(f"   平均 {sum(nums) / len(nums) * 100:.2f}%｜最低 {min(nums) * 100:g}%｜最高 {max(nums) * 100:g}%"
              f"｜低於 80%：{low}｜100% 或以上：{high}")
    print("   沒有數字的格：", dict(marks))
    for key, label in (("cat", "產品種類"), ("type", "紅利類別"), ("cur", "貨幣")):
        print(f"   {label}：", dict(collections.Counter(r[key] for r in rows)))


def main():
    bad = 0
    for who, url in PAGES:
        print(f"=== {who} ===")
        link = find_pdf(fetch(url).decode("utf-8", errors="replace"), url)
        if not link:
            bad += 1
            print("✗ 頁面內找不到「實現率」PDF 連結")
            continue
        print(f"   {link}")
        rows, b = parse_pdf(fetch(link))
        bad += b
        report(rows)
        print()
    print(f"試讀完成（只讀，沒有上傳）｜需要留意的地方：{bad} 處")


if __name__ == "__main__":
    main()
