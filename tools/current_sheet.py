"""現行の収支一覧表(期ごとのシート、得意先×月の 金額/入金日/実入金額)を読むための共通処理。

レイアウト:
  - 1行目 AK 列に年度(例: 2024 → 2024年10月〜2025年9月)
  - 収入: 5行目から「収入合計」行の手前まで。B=得意先
  - 支出: 「支出」見出しの2行下から「支出合計」行の手前まで。B=勘定, C=支出先
  - D列から3列ずつ 10月〜9月(金額 / 入金日・支払日 / 実入金額・手数料)、AN=合計、AO=実入金額合計、AP=備考
"""
from collections import Counter, defaultdict
from datetime import date, datetime

from openpyxl import load_workbook

FEE_TOLERANCE = 1000  # これ以下の差額は振込手数料の差引とみなす


def month_of(fy, m):
    """m=0..11 (10月始まり) → (年, 月)"""
    mon = (m + 9) % 12 + 1
    return (fy if mon >= 10 else fy + 1), mon


def num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def blank(x):
    return x is None or (isinstance(x, str) and x.strip() == "")


def ymd(x):
    if isinstance(x, datetime):
        return x.date()
    return x if isinstance(x, date) else None


def load(path):
    """(数式のブック, 値のブック, [(シート名, 年度, 区分の行範囲)]) を返す"""
    wb_f = load_workbook(path)
    wb_v = load_workbook(path, data_only=True)
    sheets = []
    for ws in wb_f.worksheets:
        fy = ws["AK1"].value
        if not isinstance(fy, int):
            continue
        total = {ws.cell(r, 2).value: r for r in range(1, ws.max_row + 1)
                 if ws.cell(r, 2).value in ("収入合計", "支出合計")}
        exp_head = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "支出")
        sections = {"収入": (5, total["収入合計"] - 1), "支出": (exp_head + 2, total["支出合計"] - 1)}
        sheets.append((ws.title, fy, sections))
    return wb_f, wb_v, sheets


def income_cells(wb_v, sheets):
    """収入欄の (シート名, 年度, 行, 得意先, m, 金額, 入金日, 実入金額) を列挙"""
    for title, fy, sections in sheets:
        ws = wb_v[title]
        r0, r1 = sections["収入"]
        for r in range(r0, r1 + 1):
            name = ws.cell(r, 2).value
            for m in range(12):
                c0 = 4 + 3 * m
                yield (title, fy, r, name, m,
                       ws.cell(r, c0).value, ws.cell(r, c0 + 1).value, ws.cell(r, c0 + 2).value)


def payment_terms(wb_v, sheets):
    """金額どおり入金された実績(手数料差引を含む)から、得意先ごとの支払条件を推定する。
    戻り値: {得意先: (支払サイト(月), 支払日(31=末日), 根拠件数, 確からしさ)}"""
    samples = defaultdict(list)
    for _, fy, _, name, m, amt, dt, paid in income_cells(wb_v, sheets):
        d = ymd(dt)
        if blank(name) or d is None or not (num(amt) and num(paid)) or abs(amt - paid) > FEE_TOLERANCE:
            continue
        y, mon = month_of(fy, m)
        lag = (d.year * 12 + d.month) - (y * 12 + mon)
        if 0 <= lag <= 6:
            samples[name.strip()].append((lag, 31 if d.day >= 25 else d.day))
    result = {}
    for name, s in samples.items():
        lag, lag_n = Counter(x[0] for x in s).most_common(1)[0]
        days = [x[1] for x in s if x[0] == lag]
        day, day_n = Counter(days).most_common(1)[0]
        if len(s) >= 3 and lag_n / len(s) >= 0.7 and day_n / len(days) >= 0.6:
            conf = "高"
        elif len(s) >= 2:
            conf = "中"
        else:
            conf = "低"
        result[name] = (lag, day, len(s), conf)
    return result


def terms_text(lag, day):
    lag_t = {0: "当月", 1: "翌月", 2: "翌々月"}.get(lag, f"{lag}か月後")
    return f"{lag_t}{'末日' if day == 31 else f'{day}日'}"


def customers(wb_v, sheets):
    """得意先名の一覧(新しい期の並び順を優先)と、各得意先の情報を返す。
    戻り値: [(得意先, {"期": [...], "請求月数": 最新の期で請求があった月数})]"""
    info = {}
    for title, fy, sections in reversed(sheets):
        ws = wb_v[title]
        r0, r1 = sections["収入"]
        for r in range(r0, r1 + 1):
            name = ws.cell(r, 2).value
            if blank(name):
                continue
            name = name.strip()
            months = sum(1 for m in range(12) if num(ws.cell(r, 4 + 3 * m).value))
            d = info.setdefault(name, {"期": [], "請求月数": {}})
            d["期"].append(title)
            d["請求月数"][title] = d["請求月数"].get(title, 0) + months
    return list(info.items())
