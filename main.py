import collections
import datetime
import io
import math
import os
import re
import statistics
from urllib.parse import urljoin

import pdfplumber
import requests
from curl_cffi import requests as cr
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

TOKEN = os.environ.get("NOTION_TOKEN", "").strip()
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": VERSION,
    "Content-Type": "application/json",
}
PAGES = [
    ("忠意人壽（香港）", "忠意人壽",
     "https://www.generali.com.hk/ZH_HK/claims_and_support/reference/generali_life"),
    ("忠意保險（香港分行）", "香港分行",
     "https://www.generali.com.hk/ZH_HK/claims_and_support/reference/generali_spa"),
]
LABELS = [str(i) for i in range(1, 11)] + ["10+"]
YEAR_RE = re.compile(r"^第(\d{1,2}\+?)個$")
PCT_RE = re.compile(r"^\d+(\.\d+)?%$")
NUM_RE = re.compile(r"^\d+(\.\d+)?$")
WIDE_RE = re.compile(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef]")
TAILS = {"停止銷售": "sales", "尚未推出": "launched", "不適用": "N/A"}
SHORT = {"停止銷售": "停售", "尚未推出": "未推出", "不適用": "不適用"}
FILLERS = {"Closed", "to", "sales", "Not", "yet", "launched", "N/A"}
SIMPLIFIED = "个实现产货红单险"
TOL = 4
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTO = "自動生成"
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"
NAME_COLOR = "FF1F2933"
CUR_COLOR = "FF7030A0"
SEP_COLOR = "FFE9ECEF"


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid(SEP_COLOR)
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color=NAME_COLOR)
CUR_FONT = Font(bold=True, color=CUR_COLOR)
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


# ---------- 讀取忠意 PDF ----------
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


def wide(ch):
    return bool(WIDE_RE.match(ch))


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
    out = []
    for w in sorted(words, key=lambda v: (v["top"], v["x0"])):
        if out and w["top"] - out[-1][0]["top"] <= 3:
            out[-1].append(w)
        else:
            out.append([w])
    return [sorted(line, key=lambda v: v["x0"]) for line in out]


def join_parts(parts):
    """中文與中文之間不加空格，其他情況加空格。"""
    out = ""
    for p in parts:
        if out and p and not (wide(out[-1]) and wide(p[0])):
            out += " "
        out += p
    return out


def line_text(line):
    return join_parts([w["text"] for w in line])


def col_lines(words):
    return [{"text": line_text(line), "top": min(w["top"] for w in line),
             "bottom": max(w["bottom"] for w in line)} for line in lines_of(words)]


def take(lines, last_bottom, h):
    """只取表格範圍內的行；最後一格的文字可以比數值多幾行。"""
    out = []
    for ln in lines:
        if ln["top"] < last_bottom or (out and ln["top"] - out[-1]["bottom"] <= 0.8 * h):
            out.append(ln)
        else:
            break
    return out


def finish(block):
    texts = block["lines"]
    last = max([i for i, t in enumerate(texts) if has_cjk(t)], default=-1)
    block["name"] = join_parts(texts[:last + 1]) or join_parts(texts)
    block["mid"] = (block["top"] + block["bottom"]) / 2


def blocks_top(lines, tops):
    """靠上對齊：新一格只會在某一行的頂部開始。"""
    out = []
    for k, ln in enumerate(lines):
        row = next((i for i, t in enumerate(tops) if abs(t - ln["top"]) <= TOL), None)
        start = not out
        if out and row is not None and row > out[-1]["row"]:
            limit = tops[row + 1] - TOL if row + 1 < len(tops) else float("inf")
            start = has_cjk(ln["text"]) or any(
                has_cjk(x["text"]) and ln["top"] < x["top"] < limit for x in lines[k + 1:])
        if start:
            out.append({"lines": [], "top": ln["top"], "bottom": ln["bottom"], "row": row or 0})
        out[-1]["lines"].append(ln["text"])
        out[-1]["bottom"] = ln["bottom"]
    return out


def blocks_centre(lines, h):
    """置中對齊：英文之後再出現中文，或者空隙較大，當作新一格。"""
    out, in_en = [], False
    for ln in lines:
        cjk = has_cjk(ln["text"])
        if not out or (cjk and in_en) or ln["top"] - out[-1]["bottom"] > 0.8 * h:
            out.append({"lines": [], "top": ln["top"], "bottom": ln["bottom"]})
        block = out[-1]
        block["lines"].append(ln["text"])
        block["bottom"] = ln["bottom"]
        in_en = any(has_cjk(t) for t in block["lines"]) and not cjk
    return out


def bounds_of(centres):
    n = len(centres)
    if n == 1:
        return [centres[0] - 18, centres[0] + 18]
    mids = [(centres[i] + centres[i + 1]) / 2 for i in range(n - 1)]
    return [2 * centres[0] - mids[0]] + mids + [2 * centres[-1] - mids[-1]]


def assign(blocks, bounds):
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


def parse_band(page, words, info, end, pno, cat):
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
    tops = [min(w["top"] for w in row) for row in grid]
    aligned = all(max(w["top"] for w in row) - min(w["top"] for w in row) <= TOL for row in grid)

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

    bounds = bounds_of(centres)
    owners, modes = [], []
    for label, lo, hi in (("產品系列", -1, b1), ("紅利類別", b1, b2), ("貨幣", b2, x_left)):
        lines = take(col_lines([w for w in zone if lo <= cx(w) < hi]), last_bottom, h)
        if not lines:
            return fail(f"「{label}」一欄沒有文字", data_top - 5, end)
        top_mode = aligned and abs(lines[0]["top"] - tops[0]) <= TOL
        if top_mode:
            blocks = blocks_top(lines, tops)
            for b in blocks:
                finish(b)
            owner = []
            for j, b in enumerate(blocks):
                nxt = blocks[j + 1]["row"] if j + 1 < len(blocks) else n
                owner += [b] * (nxt - b["row"])
        else:
            blocks = blocks_centre(lines, h)
            for b in blocks:
                finish(b)
            owner, worst = assign(blocks, bounds)
            if owner is None:
                return fail(f"「{label}」一欄讀到 {len(blocks)} 格，但數值有 {n} 行", data_top - 5, end)
            if worst > 0.8 * h:
                bad += 1
                print(f"   ⚠ 第 {pno} 頁：「{label}」與行的位置配對偏差 {worst:.0f}pt，可能配錯")
        if label == "紅利類別" and len(blocks) != n:
            return fail(f"「紅利類別」讀到 {len(blocks)} 格，但數值有 {n} 行", data_top - 5, end)
        owners.append(owner)
        modes.append("靠上" if top_mode else "置中")

    n_prod = len({id(b) for b in owners[0]})
    print(f"   第 {pno} 頁｜{cat or '（未讀到種類）'}｜{n} 行｜產品 {n_prod} 格｜對齊：{'/'.join(modes)}")
    rows = []
    for i in range(n):
        vals = [round(float(w["text"][:-1]) / 100, 4) if PCT_RE.match(w["text"]) else w["text"]
                for w in grid[i]]
        rows.append({"cat": cat, "prod": owners[0][i]["name"], "type": owners[1][i]["name"],
                     "cur": owners[2][i]["name"], "vals": vals, "page": pno})
    if bad:
        print("   --- 左邊三欄的位置（供核對） ---")
        dump([w for w in zone if cx(w) < x_left], data_top - 1, last_bottom + 40, x_left, 60)
    return rows, bad


def parse_pdf(data):
    rows, bad, seen_years = [], 0, set()
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        meta = (pdf.metadata or {}).get("CreationDate", "") or ""
        texts = [p.extract_text() or "" for p in pdf.pages]
        m = re.search(r"(\d{4})\s*報告年度", "\n".join(texts))
        if not m:
            raise SystemExit("✗ PDF 內找不到報告年度")
        year = m.group(1)
        c = re.search(r"D:(\d{4})(\d{2})(\d{2})", meta)
        created = "-".join(c.groups()) if c else "未知"
        print(f"   PDF {len(data) // 1024} KB｜{len(pdf.pages)} 頁｜報告年度 {year}｜建立 {created}")
        cat, skipped, quiet = "", [], []
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
                else:
                    quiet.append(pno)
                continue
            seen_years |= {int(y) for y in re.findall(r"[(（]\s*(\d{4})\s*年", texts[pno - 1])}
            infos = [head_info(words, b) for b in bands]
            for i, info in enumerate(infos):
                end = infos[i + 1]["start"] - 2 if i + 1 < len(infos) else page.height
                cat = info["cat"] or cat
                got, b = parse_band(page, words, info, end, pno, cat)
                rows += got
                bad += b
        print(f"   沒有表格的頁：{quiet}｜略過簡體頁：{skipped}")
    missing = sorted(set(range(int(year) - 10, int(year))) - seen_years)
    if missing:
        bad += 1
        print(f"   ✗ 表頭找不到這些保單生效年份：{missing}，版面可能改了")
    return year, created, rows, bad


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


# ---------- 平均值 ----------
def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows, n_fixed):
    return [v for r in rows for v in r[n_fixed:] if is_num(v)]


def build_summary(n_fixed, labels, spans, groups):
    width = n_fixed + len(labels)
    head = ["所有年份平均值"] + [f"持滿{k}年平均值" for k in labels[1:]]
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(n_fixed, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": list(spans), "year": [year_line]}


def fmt_pct(v):
    return f"{v * 100:.2f}%" if is_num(v) else "/"


# ---------- 生成 Excel（保誠格式） ----------
def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    """100% 或以上綠底；低於 80% 紅底；沒有數字用淺灰字。"""
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def cat_font(cat):
    s = str(cat)
    if "危疾" in s:
        return Font(color="FFC55A11")
    if "年金" in s:
        return Font(color="FF2E7D32")
    return Font(color="FF2F5597")


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 12, 10, 14]
    caps = [52, 24, 12, 44]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(n_fixed)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 66

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = cat_font(v)
            elif c == 3:
                cell.font = CUR_FONT
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = "0%"
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = n_fixed
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def add_overview(wb, year, today, items, remarks):
    """items：[(簡稱, 工作表名稱, 資料行, 固定欄數)]"""
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    def bullets(r, lines):
        for t in lines:
            r += 1
            cell = ws.cell(row=r, column=1, value="・" + t)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell.font = NOTE_FONT
            cell.alignment = LEFT_FIT
            ws.row_dimensions[r].height = ROW_H
        return r

    banner(1, f"忠意分紅實現率概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：忠意香港網站的分紅實現率 PDF｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品／版本數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    for _, title, rows, n_fixed in items:
        nums = numbers(rows, n_fixed)
        low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
        vals = [title, len({row[0] for row in rows}), len(rows), len(nums), mean_or_slash(nums), low]
        r += 1
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
        link = ws.cell(row=r, column=1)
        link.font = LINK_FONT
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{title}'!A1")
        ws.cell(row=r, column=5).number_format = "0.00%"
        heat(ws.cell(row=r, column=5), vals[4], WHITE)
        ws.cell(row=r, column=6).number_format = "0.0%"
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "按產品種類的整體平均值")
    r += 1
    table_head(r, ["產品種類"] + [short for short, _, _, _ in items])
    cats = []
    for _, _, rows, _ in items:
        for row in rows:
            if row[1] not in cats:
                cats.append(row[1])
    for cat in cats:
        r += 1
        cell = ws.cell(row=r, column=1, value=cat or "（未分類）")
        cell.border = BORDER
        cell.alignment = LEFT_FIT
        cell.font = Font(bold=True, color=cat_font(cat).color)
        for k, (_, _, rows, n_fixed) in enumerate(items):
            v = mean_or_slash(numbers([row for row in rows if row[1] == cat], n_fixed))
            c2 = ws.cell(row=r, column=k + 2, value=v)
            c2.border = BORDER
            c2.alignment = CENTER_FIT
            c2.number_format = "0.00%"
            heat(c2, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方 PDF 顯示「停止銷售」、「尚未推出」或「不適用」")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "使用提示")
    r = bullets(r, ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
                    "工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
                    "向右或向下捲動時，標題列和產品名稱會固定不動。",
                    "平均值摘要在工作表最底部。",
                    "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"])
    if remarks:
        r += 2
        section(r, "資料留意")
        bullets(r, remarks)


# ---------- Notion（沿用保誠） ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def target_id():
    raw = (os.environ.get("NOTION_DATABASE_ID") or os.environ.get("NOTION_DATABASE_URL") or "").strip()
    return raw.split("?")[0].rstrip("/").replace("-", "")[-32:]


def resolve_source(raw):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=HEADERS, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    print(f"找不到忠意的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
          "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")
    raise SystemExit(1)


def query_pages(source):
    r = requests.post(f"{API}/data_sources/{source}/query", headers=HEADERS, json={"page_size": 100}, timeout=30)
    return check(r, "讀取年份資料庫").get("results", [])


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=HEADERS, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            out.append((block["id"], block["file"].get("name", "")))
    return out


def upload_file(page_id, filename, content):
    old_auto = [bid for bid, name in file_blocks(page_id) if AUTO in name]
    r = requests.post(f"{API}/file_uploads", headers=HEADERS,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                      headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": VERSION},
                      files={"file": (filename, content, XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id}}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=HEADERS, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=HEADERS, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 主流程 ----------
def main():
    print("=== 1. 讀取官方 PDF ===")
    results, bad = [], 0
    for who, short, url in PAGES:
        print(f"--- {who} ---")
        link = find_pdf(fetch(url).decode("utf-8", errors="replace"), url)
        if not link:
            raise SystemExit("✗ 頁面內找不到「實現率」PDF 連結")
        print(f"   {link}")
        year, created, parsed, b = parse_pdf(fetch(link))
        bad += b
        report(parsed)
        results.append({"who": who, "short": short, "year": year, "created": created, "parsed": parsed})
        print()

    print("=== 2. 核對 ===")
    if bad:
        raise SystemExit(f"✗ 有 {bad} 處讀取問題，暫不生成 Excel，也不上傳")
    years = sorted({x["year"] for x in results})
    if len(years) != 1:
        raise SystemExit(f"✗ 兩份 PDF 的報告年度不一致：{years}，暫不更新")
    if any(not x["parsed"] for x in results):
        raise SystemExit("✗ 有 PDF 沒有讀到任何資料，暫不更新")
    year = years[0]
    y = int(year)
    spans = [str(y - i) for i in range(1, 11)] + [f"{y - 10}前"]
    print(f"   報告年度 {year}｜保單年度欄：" + " ".join(f"{k}({s})" for k, s in zip(LABELS, spans)))
    for x in results:
        x["rows"] = [[r["prod"], r["cat"], r["cur"], r["type"]]
                     + [v if isinstance(v, float) else "/" for v in r["vals"]] for r in x["parsed"]]
        keys = [(r[0], r[1], r[2], r[3]) for r in x["rows"]]
        if len(set(keys)) != len(keys):
            raise SystemExit(f"✗ {x['who']} 有重複的產品＋紅利類別，暫不更新")
        nums = numbers(x["rows"], 4)
        print(f"   {x['who']}｜產品 {len({(r[1], r[0]) for r in x['rows']})} 個｜{len(x['rows'])} 行"
              f"｜數值 {len(nums)} 個｜平均 {fmt_pct(mean_or_slash(nums))}｜PDF 建立 {x['created']}")

    print()
    print("=== 3. Excel ===")
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    header = ["产品名称", "产品种类", "货币", "红利类型"] + [f"保单年度{k}\n{s}" for k, s in zip(LABELS, spans)]
    wb = Workbook()
    wb.remove(wb.active)
    items = []
    for x, tab in zip(results, ("2E75B6", "C55A11")):
        rows = x["rows"]
        groups = [("所有分紅", rows)]
        for t in dict.fromkeys(r[3] for r in rows):
            sub = [r for r in rows if r[3] == t]
            if numbers(sub, 4):
                groups.append((t, sub))
        summary = build_summary(4, LABELS, spans, groups)
        notes = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時所示的金額。"
                 "100% 即與銷售時一致，高於 100% 即比銷售時多。\n"
                 "綠色＝100% 或以上；紅色＝低於 80%；「/」＝官方 PDF 顯示「停止銷售」、「尚未推出」或「不適用」。\n"
                 f"資料來源：{x['who']}的分紅實現率 PDF（報告年度 {year}，PDF 建立 {x['created']}）｜更新日期：{today}")
        title = f"{year}{x['short']}分紅實現率"
        add_sheet(wb, title, f"{x['who']}分紅實現率（報告年度 {year}）", notes, header, rows, summary,
                  "分紅類型", 4, tab)
        items.append((x["short"], title, rows, 4))
    remarks = ["「/」代表官方 PDF 顯示「停止銷售」、「尚未推出」或「不適用」。",
               "忠意人壽（香港）有限公司與忠意保險有限公司（香港分行）各自披露，所以分成兩個工作表。",
               "產品名稱和產品種類按官方 PDF 表格內的寫法。",
               "只包括分紅實現率，不包括萬用壽險的派息率。"]
    both = {r[0] for r in results[0]["rows"]} & {r[0] for r in results[1]["rows"]}
    if both:
        remarks.insert(2, "兩份 PDF 都有列出的產品：" + "、".join(sorted(both)) + "。")
    add_overview(wb, year, today, items, remarks)
    buf = io.BytesIO()
    wb.save(buf)
    content = buf.getvalue()
    path = f"generali_fulfillment_{year}.xlsx"
    with open(path, "wb") as f:
        f.write(content)
    print(f"   {path}｜{len(content) // 1024} KB｜工作表：", "、".join(wb.sheetnames))

    print()
    print("=== 4. 上傳到 Notion ===")
    if os.environ.get("UPLOAD") != "1":
        print("   試跑：沒有上傳")
        print("完成")
        return
    target = target_id()
    if not TOKEN or not target:
        raise SystemExit("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    source = resolve_source(target)
    page = next((p for p in query_pages(source) if year in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": year}}]}}}
        page = check(requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30), "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    upload_file(page["id"], f"忠意分紅實現率（{year}年）{AUTO}.xlsx", content)
    print("完成")


if __name__ == "__main__":
    main()
