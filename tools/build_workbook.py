"""請求・入金管理表 (請求入金管理表.xlsx) を生成するスクリプト。

使い方:  python tools/build_workbook.py [出力パス]
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "請求入金管理表.xlsx"

FONT = "Meiryo UI"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_NOTE = Font(name=FONT, size=9, color="595959")

FILL_HEAD = PatternFill("solid", fgColor="1F4E78")
FILL_HEAD_AUTO = PatternFill("solid", fgColor="595959")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")  # 入力欄
FILL_AUTO = PatternFill("solid", fgColor="EDEDED")  # 自動計算
FILL_SECTION = PatternFill("solid", fgColor="DDEBF7")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="top", wrap_text=True)

FMT_DATE = "yyyy/mm/dd"
FMT_MONTH = "yyyy年m月"
FMT_YEN = "#,##0;[Red]-#,##0"

# データ行の範囲
HDR = 4
FIRST = 5
INV_N, DEP_N, CLI_N = 500, 1000, 200
INV_LAST = FIRST + INV_N - 1
DEP_LAST = FIRST + DEP_N - 1
CLI_LAST = FIRST + CLI_N - 1

S_SET, S_INV, S_DEP, S_CLI = "'設定'", "'請求台帳'", "'入金明細'", "'取引先マスタ'"
S_DASH, S_BAL = "'チェック'", "'取引先別残高'"
BASE_DATE = f"{S_SET}!$C$5"
NEAR_DAYS = f"{S_SET}!$C$6"
PREFIX = f"{S_SET}!$C$4"
TARGET_MONTH = f"{S_DASH}!$C$4"


def rng(sheet, col, first=FIRST, last=None):
    last = last or {S_INV: INV_LAST, S_DEP: DEP_LAST, S_CLI: CLI_LAST, S_BAL: CLI_LAST}[sheet]
    return f"{sheet}!${col}${first}:${col}${last}"


INV = {c: rng(S_INV, c) for c in "ABCDEFGHIJKLMNOPQRSTUVWXY"}
DEP = {c: rng(S_DEP, c) for c in "ABCDEFGHIJK"}
CLI = {c: rng(S_CLI, c) for c in "ABCDEFGHIJKL"}

wb = Workbook()


def style_range(ws, ref, font=F_BASE, fill=None, fmt=None, align=None, border=True):
    for row in ws[ref]:
        for c in row:
            c.font = font
            if fill:
                c.fill = fill
            if fmt:
                c.number_format = fmt
            if align:
                c.alignment = align
            if border:
                c.border = BORDER


def title(ws, text, note):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = note
    ws["A2"].font = F_NOTE
    ws.sheet_view.showGridLines = False


def table(ws, columns, first, last):
    """columns: (見出し, 幅, 'in'|'auto', 書式, 数式テンプレート or None, コメント)"""
    for i, (head, width, kind, fmt, formula, comment) in enumerate(columns, start=1):
        col = get_column_letter(i)
        cell = ws.cell(row=HDR, column=i, value=head)
        cell.font = F_HEAD
        cell.fill = FILL_HEAD if kind == "in" else FILL_HEAD_AUTO
        cell.alignment = CENTER
        cell.border = BORDER
        if comment:
            cell.comment = Comment(comment, "管理表")
        ws.column_dimensions[col].width = width
        fill = FILL_INPUT if kind == "in" else FILL_AUTO
        for r in range(first, last + 1):
            c = ws.cell(row=r, column=i)
            c.font = F_BASE
            c.fill = fill
            c.border = BORDER
            if fmt:
                c.number_format = fmt
            if formula:
                c.value = formula.format(r=r)
            if kind == "in":
                c.protection = Protection(locked=False)
    ws.row_dimensions[HDR].height = 32
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.auto_filter.ref = f"A{HDR}:{get_column_letter(len(columns))}{last}"


def protect(ws):
    # パスワードなし。数式の誤上書きを防ぐ目的(フィルタは使用可、並べ替えは不可)
    p = ws.protection
    p.sheet = True
    p.autoFilter = False
    p.formatColumns = False
    p.formatRows = False
    p.formatCells = False


def add_list(ws, ref, source, allow_blank=True, error=None):
    dv = DataValidation(type="list", formula1=source, allow_blank=allow_blank)
    dv.error = error or "一覧から選択してください"
    dv.errorTitle = "入力エラー"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(ref)


def add_date(ws, ref):
    dv = DataValidation(type="date", operator="between",
                        formula1="DATE(2000,1,1)", formula2="DATE(2100,12,31)", allow_blank=True)
    dv.error = "日付を入力してください(例: 2026/9/30)"
    dv.errorTitle = "入力エラー"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(ref)


def add_number(ws, ref, minimum=0):
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1=str(minimum),
                        allow_blank=True)
    dv.error = "数値(円)を入力してください"
    dv.errorTitle = "入力エラー"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(ref)


def status_colors(ws, ref, col_letter):
    rules = [
        ("入金済", "E2EFDA", "375623"),
        ("期日超過", "F8CBAD", "9C0006"),
        ("一部入金・超過", "F8CBAD", "9C0006"),
        ("未送付", "FCE4D6", "C65911"),
        ("要入力", "FFFF00", "9C0006"),
        ("期日間近", "FFE699", "7F6000"),
        ("一部入金", "FFE699", "7F6000"),
        ("過入金", "E4DFEC", "5B2C6F"),
        ("取消", "D9D9D9", "7F7F7F"),
        ("未消込", "F8CBAD", "9C0006"),
        ("番号誤り", "FFFF00", "9C0006"),
        ("消込済", "E2EFDA", "375623"),
    ]
    first = ref.split(":")[0]
    row = "".join(ch for ch in first if ch.isdigit())
    for text, bg, fg in rules:
        ws.conditional_formatting.add(ref, FormulaRule(
            formula=[f'${col_letter}{row}="{text}"'],
            fill=PatternFill("solid", fgColor=bg, bgColor=bg),
            font=Font(name=FONT, bold=True, color=fg)))


# ---------------------------------------------------------------- はじめに
ws = wb.active
ws.title = "はじめに"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 90
lines = [
    ("title", "請求・入金管理表  ― 使い方"),
    ("note", "請求書の「発行漏れ」「送付漏れ」「入金確認漏れ」「記入漏れ」を仕組みで防ぐための管理表です。"),
    ("", ""),
    ("section", "1. 色の見方"),
    ("legend_in", "黄色のセル", "入力するセル(ここだけ入力します)"),
    ("legend_auto", "灰色のセル", "自動計算(入力不要。シート保護で上書きできないようにしています)"),
    ("", ""),
    ("section", "2. シート構成"),
    ("row", "請求台帳", "請求書1枚=1行。請求番号は台帳が自動採番します。ステータス・入力漏れを自動判定。"),
    ("row", "入金明細", "通帳/ネットバンキングの入金を【全件】転記し、消込先の請求番号を入れます。"),
    ("row", "取引先マスタ", "取引先コード・支払条件(サイト/支払日)・毎月請求の有無を登録します。"),
    ("row", "チェック", "未送付/期日超過/未消込/要入力などの件数を一覧表示。0件でない項目が「要対応」です。"),
    ("row", "取引先別残高", "取引先ごとの未回収残高と超過日数の内訳、毎月請求先の請求漏れ候補を表示します。"),
    ("row", "月次チェック記録", "月次締めのチェック結果と承認を記録します(誰がいつ確認したかを残す)。"),
    ("row", "設定", "請求番号の接頭辞、基準日、期日間近の日数、プルダウンの選択肢。"),
    ("", ""),
    ("section", "3. 基本の流れ"),
    ("row", "① 請求書を作る前", "請求台帳に1行追加(請求日・取引先・件名・金額・税率・送付方法)→ 採番された請求番号を請求書に記載。"),
    ("row", "② 請求書を送ったら", "送付日・送付者を入力。未入力の間はステータスが「未送付」のまま残ります。"),
    ("row", "③ 入金があったら", "入金明細に入金日・名義・金額を転記し、消込先の請求番号を入力。台帳の入金額・残高が自動更新されます。"),
    ("row", "④ 消込を確認したら", "残高が0になった請求は、別の人が確認して「消込確認者」に名前を入れます(未入力だと「要入力」)。"),
    ("row", "⑤ 毎週", "チェックシートで「期日超過」「期日間近」を確認し、担当者から督促・状況確認。"),
    ("row", "⑥ 毎月(締め後)", "チェックシートの全項目が「OK」になることを確認し、月次チェック記録に実施者・承認者を記入。"),
    ("", ""),
    ("section", "4. ステータスの意味(請求台帳)"),
    ("row", "要入力", "必須項目の漏れ、取引先コード誤り、送付者・消込確認者の未入力がある。「入力漏れ」列に項目名を表示。"),
    ("row", "未送付", "台帳登録済みだが送付日が空欄。請求書の出し忘れ・送り忘れの可能性。"),
    ("row", "入金待ち / 期日間近", "送付済みで期日前。期日まで「設定」の日数以内なら期日間近。"),
    ("row", "期日超過", "入金期日を過ぎても未回収残高がある。督促が必要。"),
    ("row", "一部入金 / 一部入金・超過", "入金額が請求額に足りない(振込手数料の差引なら「手数料・値引等」に入力)。"),
    ("row", "過入金", "請求額より多く入金されている(二重振込・消込先誤りの可能性)。"),
    ("row", "入金済", "未回収残高が0。"),
    ("row", "取消", "「取消」列で取消した請求。行は削除せず残します(番号の欠番を作らないため)。"),
    ("", ""),
    ("section", "5. 守るルール"),
    ("row", "行を削除しない", "間違えた請求は「取消」を選び、備考に理由を書く。番号の連続性が発行漏れチェックの根拠になります。"),
    ("row", "並べ替えをしない", "請求番号は行位置から採番されるため、並べ替えは不可(シート保護で制限)。絞り込みはフィルタで。"),
    ("row", "入金は全件転記", "請求に対応しない入金も含めて全件入力。「未消込」が残っていれば消込漏れか不明入金です。"),
    ("row", "1入金で複数請求", "入金明細を請求の数だけ行に分け、金額を按分して各請求番号を入力。"),
    ("row", "振込手数料の差引", "先方負担の手数料が差し引かれた場合、請求台帳の「手数料・値引等」に差額を入力。"),
    ("row", "サンプルデータ", "各シートの「サンプル」「SAMPLE」行は動作確認用です。運用開始前に入力セルの内容を消去してください。"),
    ("row", "シート保護", "パスワードなしで保護しています。列の追加などが必要な場合は[校閲]→[シート保護の解除]。"),
]
r = 1
for item in lines:
    kind = item[0]
    if kind == "title":
        ws.cell(row=r, column=2, value=item[1]).font = F_TITLE
    elif kind == "note":
        ws.cell(row=r, column=2, value=item[1]).font = F_NOTE
    elif kind == "section":
        c = ws.cell(row=r, column=2, value=item[1])
        c.font = F_BOLD
        for col in (2, 3):
            ws.cell(row=r, column=col).fill = FILL_SECTION
    elif kind in ("row", "legend_in", "legend_auto"):
        a = ws.cell(row=r, column=2, value=item[1])
        b = ws.cell(row=r, column=3, value=item[2])
        a.font, b.font = F_BOLD, F_BASE
        a.alignment, b.alignment = WRAP, WRAP
        a.border = b.border = BORDER
        if kind == "legend_in":
            a.fill = FILL_INPUT
        if kind == "legend_auto":
            a.fill = FILL_AUTO
    r += 1

# ---------------------------------------------------------------- 設定
st = wb.create_sheet("設定")
title(st, "設定", "黄色のセルを必要に応じて変更してください。プルダウンの選択肢は各リストに追記できます。")
st.column_dimensions["A"].width = 3
st.column_dimensions["B"].width = 22
st.column_dimensions["C"].width = 16
st.column_dimensions["D"].width = 50
settings = [
    (4, "請求番号の接頭辞", "INV-", None, "請求番号 = 接頭辞 + 5桁の連番(例: INV-00001)"),
    (5, "基準日", "=TODAY()", FMT_DATE, "通常は =TODAY() のまま。過去時点で確認したい場合のみ日付を直接入力"),
    (6, "期日間近とする日数", 7, "0\"日\"", "入金期日まで何日以内を「期日間近」とするか"),
]
for row, label, val, fmt, note in settings:
    st.cell(row=row, column=2, value=label).font = F_BOLD
    c = st.cell(row=row, column=3, value=val)
    c.font, c.fill, c.border = F_BASE, FILL_INPUT, BORDER
    if fmt:
        c.number_format = fmt
    st.cell(row=row, column=4, value=note).font = F_NOTE

lists = {
    "F": ("税率", [0.10, 0.08, 0.0], "0%"),
    "G": ("送付方法", ["メール(PDF)", "郵送", "請求書発行システム", "手渡し", "その他"], None),
    "H": ("入金口座", ["○○銀行 本店 普通", "△△信用金庫 普通"], None),
    "I": ("取消", ["取消"], None),
    "J": ("有無", ["有", "無"], None),
    "K": ("実施", ["済", "該当なし"], None),
}
LIST_FIRST, LIST_LAST = 5, 24
for col, (head, values, fmt) in lists.items():
    st.column_dimensions[col].width = 20
    h = st[f"{col}4"]
    h.value, h.font, h.fill, h.alignment = head, F_HEAD, FILL_HEAD, CENTER
    for i in range(LIST_FIRST, LIST_LAST + 1):
        c = st[f"{col}{i}"]
        c.font, c.fill, c.border = F_BASE, FILL_INPUT, BORDER
        if fmt:
            c.number_format = fmt
    for i, v in enumerate(values):
        st[f"{col}{LIST_FIRST + i}"] = v
st["F2"] = "▼ プルダウンの選択肢(空欄に追記すると選択肢に増えます)"
st["F2"].font = F_NOTE


def list_src(col):
    return f"{S_SET}!${col}${LIST_FIRST}:${col}${LIST_LAST}"


# ---------------------------------------------------------------- 取引先マスタ
cm = wb.create_sheet("取引先マスタ")
title(cm, "取引先マスタ", "取引先ごとの支払条件を登録します。請求台帳の入金期日は「支払サイト」「支払日」から自動計算されます。")
cli_cols = [
    ("取引先コード", 12, "in", "@", None, "重複不可。請求台帳ではこのコードで選択します"),
    ("取引先名", 28, "in", None, None, None),
    ("振込名義(カナ)", 22, "in", None, None, "通帳に表示される名義。入金明細での照合に使います"),
    ("締日(参考)", 10, "in", None, None, "例: 末日 / 20日"),
    ("支払サイト(月)", 11, "in", "0", None, "請求月の何か月後に支払われるか。翌月=1、翌々月=2、当月=0"),
    ("支払日", 9, "in", "0", None, "1〜31。末日払いは31(月の日数に合わせて自動調整)"),
    ("毎月請求", 9, "in", None, None, "毎月必ず請求する取引先は「有」。当月の請求が未登録だと「請求漏れ候補」になります"),
    ("既定の送付方法", 16, "in", None, None, None),
    ("送付先(メール/住所)", 30, "in", None, None, None),
    ("営業担当", 12, "in", None, None, "督促・状況確認の担当者"),
    ("備考", 30, "in", None, None, None),
    ("重複チェック", 12, "auto", None,
     '=IF(A{r}="","",IF(COUNTIF($A$5:$A$' + str(CLI_LAST) + ',A{r})>1,"コード重複",""))', None),
]
table(cm, cli_cols, FIRST, CLI_LAST)
add_list(cm, f"G{FIRST}:G{CLI_LAST}", list_src("J"))
add_list(cm, f"H{FIRST}:H{CLI_LAST}", list_src("G"))
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="12", allow_blank=True)
dv.error, dv.showErrorMessage = "0〜12の整数を入力してください", True
cm.add_data_validation(dv)
dv.add(f"E{FIRST}:E{CLI_LAST}")
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="31", allow_blank=True)
dv.error, dv.showErrorMessage = "1〜31の整数を入力してください(末日=31)", True
cm.add_data_validation(dv)
dv.add(f"F{FIRST}:F{CLI_LAST}")
cm.conditional_formatting.add(f"L{FIRST}:L{CLI_LAST}", FormulaRule(
    formula=[f'$L{FIRST}<>""'], fill=PatternFill("solid", fgColor="FFFF00", bgColor="FFFF00"),
    font=Font(name=FONT, bold=True, color="9C0006")))
sample_clients = [
    ("C001", "株式会社サンプル商事", "カ)サンプルシヨウジ", "末日", 1, 31, "有", "メール(PDF)",
     "keiri@sample.example", "山田", "サンプル(運用前に削除)"),
    ("C002", "有限会社テスト工業", "ユ)テストコウギヨウ", "末日", 2, 10, "無", "郵送",
     "東京都○○区…", "佐藤", "サンプル(運用前に削除)"),
    ("C003", "合同会社デモ", "ド)デモ", "20日", 1, 31, "有", "請求書発行システム",
     "", "山田", "サンプル(運用前に削除)"),
]
for i, rowvals in enumerate(sample_clients):
    for j, v in enumerate(rowvals, start=1):
        cm.cell(row=FIRST + i, column=j, value=v)
protect(cm)

# ---------------------------------------------------------------- 請求台帳
iv = wb.create_sheet("請求台帳")
title(iv, "請求台帳",
      "請求書を作る前にここへ登録 → 採番された請求番号を請求書に記載。黄色=入力 / 灰色=自動。行は削除せず、誤りは「取消」を選択。")
def cli_lookup(col):
    return "INDEX(" + CLI[col] + ",MATCH($C{r}," + CLI["A"] + ",0))"


due_auto = ('=IF(OR(B{r}="",C{r}="",D{r}="コード誤り"),"",IFERROR('
            'DATE(YEAR(B{r}),MONTH(B{r})+' + cli_lookup("E") + ',MIN(' + cli_lookup("F") +
            ',DAY(EOMONTH(B{r},' + cli_lookup("E") + ')))),""))')
missing = ('=IF(A{r}="","",TRIM('
           'IF(B{r}="","請求日 ","")'
           '&IF(C{r}="","取引先 ",IF(D{r}="コード誤り","取引先コード誤り ",""))'
           '&IF(E{r}="","件名 ","")'
           '&IF(F{r}="","税抜金額 ","")'
           '&IF(G{r}="","税率 ","")'
           '&IF(M{r}="","送付方法 ","")'
           '&IF(AND(N{r}<>"",O{r}=""),"送付者 ","")'
           '&IF(AND(I{r}<>"",L{r}=""),"入金期日 ","")'
           '&IF(AND(N{r}<>"",P{r}>0,S{r}<=0,X{r}=""),"消込確認者 ","")))')
status = ('=IF(A{r}="","",IF(T{r}="取消","取消",IF(W{r}<>"","要入力",IF(N{r}="","未送付",'
          'IF(S{r}=0,"入金済",IF(S{r}<0,"過入金",'
          'IF(' + BASE_DATE + '>L{r},IF(P{r}>0,"一部入金・超過","期日超過"),'
          'IF(P{r}>0,"一部入金",IF(L{r}-' + BASE_DATE + '<=' + NEAR_DAYS +
          ',"期日間近","入金待ち")))))))))')
inv_cols = [
    ("請求番号", 12, "auto", None,
     '=IF(COUNTA(B{r},C{r},E{r},F{r})=0,"",' + PREFIX + '&TEXT(ROW()-' + str(FIRST - 1) + ',"00000"))',
     "請求日・取引先・件名・金額のいずれかを入力すると自動採番。この番号を請求書に記載します"),
    ("請求日", 11, "in", FMT_DATE, None, "締日ベースの請求日。入金期日の起算月になります"),
    ("取引先コード", 10, "in", "@", None, "取引先マスタのコードから選択"),
    ("取引先名", 24, "auto", None,
     '=IF(C{r}="","",IFERROR(' + cli_lookup("B") + ',"コード誤り"))', None),
    ("件名", 28, "in", None, None, "例: 2026年9月分 保守料"),
    ("税抜金額", 12, "in", FMT_YEN, None, None),
    ("税率", 7, "in", "0%", None, None),
    ("消費税", 11, "auto", FMT_YEN, '=IF(OR(F{r}="",G{r}=""),"",ROUNDDOWN(F{r}*G{r},0))',
     "税抜金額×税率(1円未満切捨て)"),
    ("請求額(税込)", 13, "auto", FMT_YEN, '=IF(OR(F{r}="",G{r}=""),"",F{r}+H{r})', None),
    ("入金期日(自動)", 12, "auto", FMT_DATE, due_auto, "請求日と取引先マスタの支払条件から自動計算"),
    ("入金期日(個別)", 12, "in", FMT_DATE, None, "個別に期日が決まっている場合のみ入力(自動より優先)"),
    ("入金期日", 12, "auto", FMT_DATE, '=IF(K{r}<>"",K{r},J{r})', "この期日でステータスを判定"),
    ("送付方法", 14, "in", None, None, None),
    ("送付日", 11, "in", FMT_DATE, None, "送付したら必ず入力。空欄の間は「未送付」"),
    ("送付者", 10, "in", None, None, None),
    ("入金額", 12, "auto", FMT_YEN,
     '=IF(A{r}="","",SUMIF(' + DEP["F"] + ',A{r},' + DEP["E"] + '))',
     "入金明細で、この請求番号に消込した金額の合計"),
    ("最終入金日", 11, "auto", FMT_DATE,
     '=IF(OR(A{r}="",N(P{r})=0),"",SUMPRODUCT(MAX((' + DEP["F"] + '=A{r})*' + DEP["B"] + ')))', None),
    ("手数料・値引等", 11, "in", FMT_YEN, None, "先方が振込手数料を差し引いた場合などの差額。入れると残高から控除"),
    ("未回収残高", 12, "auto", FMT_YEN, '=IF(OR(A{r}="",I{r}=""),"",I{r}-N(P{r})-N(R{r}))', None),
    ("取消", 7, "in", None, None, "誤発行・再発行などで無効にする請求は「取消」を選び、備考に理由を記入(行は削除しない)"),
    ("ステータス", 13, "auto", None, status, "「はじめに」シートの4.を参照"),
    ("超過日数", 8, "auto", "0", '=IF(OR(U{r}="期日超過",U{r}="一部入金・超過"),' + BASE_DATE + '-L{r},"")', None),
    ("入力漏れ", 24, "auto", None, missing, "未入力の項目名を表示。空欄になればOK"),
    ("消込確認者", 10, "in", None, None, "入金済になった請求を確認した人(入金明細の入力者とは別の人が望ましい)"),
    ("備考", 30, "in", None, None, None),
]
table(iv, inv_cols, FIRST, INV_LAST)
iv.freeze_panes = "B5"
add_date(iv, f"B{FIRST}:B{INV_LAST}")
add_date(iv, f"K{FIRST}:K{INV_LAST}")
add_date(iv, f"N{FIRST}:N{INV_LAST}")
add_list(iv, f"C{FIRST}:C{INV_LAST}", rng(S_CLI, "A"),
         error="取引先マスタに登録されたコードを選択してください")
add_number(iv, f"F{FIRST}:F{INV_LAST}", minimum=-99999999999)
add_number(iv, f"R{FIRST}:R{INV_LAST}", minimum=-99999999999)
add_list(iv, f"G{FIRST}:G{INV_LAST}", list_src("F"))
add_list(iv, f"M{FIRST}:M{INV_LAST}", list_src("G"))
add_list(iv, f"T{FIRST}:T{INV_LAST}", list_src("I"))
status_colors(iv, f"U{FIRST}:U{INV_LAST}", "U")
iv.conditional_formatting.add(f"W{FIRST}:W{INV_LAST}", FormulaRule(
    formula=[f'$W{FIRST}<>""'], fill=PatternFill("solid", fgColor="FFFF00", bgColor="FFFF00"),
    font=Font(name=FONT, bold=True, color="9C0006")))
iv.conditional_formatting.add(f"A{FIRST}:T{INV_LAST}", FormulaRule(
    formula=[f'$T{FIRST}="取消"'], font=Font(name=FONT, strike=True, color="7F7F7F")))

today = date(2026, 9, 29)
sample_inv = [
    # 請求日, コード, 件名, 税抜, 税率, 個別期日, 送付方法, 送付日, 送付者, 手数料, 取消, 確認者, 備考
    (date(2026, 7, 31), "C001", "2026年7月分 保守料", 100000, 0.10, None, "メール(PDF)", date(2026, 7, 31),
     "鈴木", None, None, "田中", "SAMPLE: 入金済"),
    (date(2026, 7, 31), "C002", "○○部品 納品分", 250000, 0.10, None, "郵送", date(2026, 8, 3),
     "鈴木", None, None, None, "SAMPLE: 手数料差引で入金(確認者未入力→要入力)"),
    (date(2026, 6, 30), "C002", "○○部品 6月納品分", 150000, 0.10, None, "郵送", date(2026, 7, 1),
     "鈴木", None, None, None, "SAMPLE: 期日超過"),
    (date(2026, 8, 31), "C003", "8月分 業務委託料", 80000, 0.10, None, "請求書発行システム", date(2026, 9, 1),
     "鈴木", None, None, None, "SAMPLE: 一部入金"),
    (date(2026, 9, 30), "C001", "2026年9月分 保守料", 100000, 0.10, None, "メール(PDF)", None,
     None, None, None, None, "SAMPLE: 未送付"),
    (date(2026, 9, 15), "C003", "追加作業費", 30000, 0.10, None, "請求書発行システム", date(2026, 9, 15),
     "鈴木", None, "取消", None, "SAMPLE: 金額誤りのため取消→次行で再発行"),
    (date(2026, 9, 15), "C003", "追加作業費(再発行)", 33000, 0.10, date(2026, 10, 5), "請求書発行システム",
     date(2026, 9, 16), "鈴木", None, None, None, "SAMPLE: 個別期日・期日間近"),
]
cols_in = ["B", "C", "E", "F", "G", "K", "M", "N", "O", "R", "T", "X", "Y"]
for i, vals in enumerate(sample_inv):
    for col, v in zip(cols_in, vals):
        if v is not None:
            iv[f"{col}{FIRST + i}"] = v
protect(iv)

# ---------------------------------------------------------------- 入金明細
dp = wb.create_sheet("入金明細")
title(dp, "入金明細",
      "通帳・ネットバンキングの入金を全件転記し、「消込先請求番号」を入力します。1つの入金で複数請求を払われた場合は行を分けて金額を按分。")
dep_state = ('=IF(COUNTA(B{r},D{r},E{r},F{r})=0,"",IF(OR(B{r}="",E{r}=""),"要入力",'
             'IF(F{r}="","未消込",IF(ISNA(MATCH(F{r},' + INV["A"] + ',0)),"番号誤り",'
             'IF(INDEX(' + INV["T"] + ',MATCH(F{r},' + INV["A"] + ',0))="取消","取消請求へ消込","消込済")))))')
dep_cols = [
    ("No", 6, "auto", None, '=IF(COUNTA(B{r},D{r},E{r},F{r})=0,"",ROW()-' + str(FIRST - 1) + ')', None),
    ("入金日", 11, "in", FMT_DATE, None, None),
    ("入金口座", 18, "in", None, None, None),
    ("振込名義", 24, "in", None, None, "通帳の表示どおりに入力"),
    ("入金額", 12, "in", FMT_YEN, None, None),
    ("消込先請求番号", 13, "in", None, None, "対応する請求台帳の請求番号。不明なら空欄のまま(=未消込として残る)"),
    ("取引先名(請求)", 24, "auto", None,
     '=IF(F{r}="","",IFERROR(INDEX(' + INV["D"] + ',MATCH(F{r},' + INV["A"] + ',0)),""))',
     "消込先請求の取引先。振込名義と一致しているか目視で確認"),
    ("請求額(税込)", 12, "auto", FMT_YEN,
     '=IF(F{r}="","",IFERROR(INDEX(' + INV["I"] + ',MATCH(F{r},' + INV["A"] + ',0)),""))', None),
    ("消込状態", 12, "auto", None, dep_state, None),
    ("記帳者", 10, "in", None, None, None),
    ("備考", 30, "in", None, None, None),
]
table(dp, dep_cols, FIRST, DEP_LAST)
add_date(dp, f"B{FIRST}:B{DEP_LAST}")
add_number(dp, f"E{FIRST}:E{DEP_LAST}", minimum=-99999999999)
add_list(dp, f"C{FIRST}:C{DEP_LAST}", list_src("H"))
status_colors(dp, f"I{FIRST}:I{DEP_LAST}", "I")
dp.conditional_formatting.add(f"I{FIRST}:I{DEP_LAST}", FormulaRule(
    formula=[f'$I{FIRST}="取消請求へ消込"'], fill=PatternFill("solid", fgColor="FFFF00", bgColor="FFFF00"),
    font=Font(name=FONT, bold=True, color="9C0006")))
sample_dep = [
    (date(2026, 8, 31), "○○銀行 本店 普通", "カ)サンプルシヨウジ", 110000, "INV-00001", "佐々木", "SAMPLE"),
    (date(2026, 9, 10), "○○銀行 本店 普通", "ユ)テストコウギヨウ", 274560, "INV-00002", "佐々木",
     "SAMPLE: 振込手数料440円差引"),
    (date(2026, 9, 30), "○○銀行 本店 普通", "ド)デモ", 50000, "INV-00004", "佐々木", "SAMPLE: 一部入金"),
    (date(2026, 9, 28), "△△信用金庫 普通", "ヤマダ タロウ", 5500, None, "佐々木", "SAMPLE: 不明入金(未消込)"),
]
for i, vals in enumerate(sample_dep):
    for col, v in zip(["B", "C", "D", "E", "F", "J", "K"], vals):
        if v is not None:
            dp[f"{col}{FIRST + i}"] = v
iv["R6"] = 440  # サンプル: INV-00002 の振込手数料差引
protect(dp)

# ---------------------------------------------------------------- チェック
ck = wb.create_sheet("チェック")
title(ck, "チェック(ダッシュボード)",
      "判定が「要対応」の項目を解消してください。月次締めでは全項目「OK」を確認して「月次チェック記録」に記入します。")
ck.column_dimensions["A"].width = 3
ck.column_dimensions["B"].width = 40
ck.column_dimensions["C"].width = 14
ck.column_dimensions["D"].width = 16
ck.column_dimensions["E"].width = 11
ck.column_dimensions["F"].width = 60
ck["B4"] = "対象月(月次チェック用)"
ck["B4"].font = F_BOLD
ck["C4"] = "=DATE(YEAR(" + BASE_DATE + "),MONTH(" + BASE_DATE + ")-1,1)"
ck["C4"].number_format = FMT_MONTH
ck["C4"].fill, ck["C4"].border, ck["C4"].font = FILL_INPUT, BORDER, F_BOLD
ck["C4"].protection = Protection(locked=False)
ck["D4"] = "既定は基準日の前月。確認したい月の1日を入力(例: 2026/9/1)"
ck["D4"].font = F_NOTE
ck["B5"] = "基準日"
ck["B5"].font = F_BOLD
ck["C5"] = "=" + BASE_DATE
ck["C5"].number_format = FMT_DATE
ck["C5"].font = F_BOLD
ck["D5"] = "「設定」シートで変更"
ck["D5"].font = F_NOTE

for i, h in enumerate(["チェック項目", "件数", "金額", "判定", "対応方法"], start=2):
    c = ck.cell(row=7, column=i, value=h)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER



def cnt(stat):
    return "COUNTIF(" + INV["U"] + ',"' + stat + '")'


def amt(stat):
    return "SUMIF(" + INV["U"] + ',"' + stat + '",' + INV["S"] + ")"


checks = [
    ("section", "■ 請求(請求台帳)"),
    ("要入力(入力漏れ)", "=" + cnt("要入力"), None, True,
     "請求台帳を「ステータス=要入力」で絞り込み、「入力漏れ」列の項目を入力"),
    ("未送付", "=" + cnt("未送付"), "=SUMIF(" + INV["U"] + ',"未送付",' + INV["I"] + ")", True,
     "請求書を送付し送付日・送付者を入力。不要な請求なら「取消」"),
    ("期日超過(一部入金含む)", "=" + cnt("期日超過") + "+" + cnt("一部入金・超過"),
     "=" + amt("期日超過") + "+" + amt("一部入金・超過"), True,
     "営業担当から先方へ確認・督促。入金済なら入金明細の転記・消込漏れを確認"),
    ("過入金", "=" + cnt("過入金"), "=" + amt("過入金"), True,
     "二重振込・消込先の誤りを確認。返金または次回請求と相殺し備考に記録"),
    ("期日間近(参考)", "=" + cnt("期日間近"), "=" + amt("期日間近"), False, "期日前の確認・リマインド(任意)"),
    ("入金待ち(参考)", "=" + cnt("入金待ち") + "+" + cnt("一部入金"),
     "=" + amt("入金待ち") + "+" + amt("一部入金"), False, "対応不要"),
    ("section", "■ 入金(入金明細)"),
    ("未消込の入金", "=COUNTIF(" + DEP["I"] + ',"未消込")', "=SUMIF(" + DEP["I"] + ',"未消込",' + DEP["E"] + ")",
     True, "振込名義・金額から該当請求を探して請求番号を入力。不明なら先方・営業に確認"),
    ("消込先の請求番号誤り", "=COUNTIF(" + DEP["I"] + ',"番号誤り")',
     "=SUMIF(" + DEP["I"] + ',"番号誤り",' + DEP["E"] + ")", True, "請求番号を正しく入力し直す"),
    ("取消した請求への消込", "=COUNTIF(" + DEP["I"] + ',"取消請求へ消込")',
     "=SUMIF(" + DEP["I"] + ',"取消請求へ消込",' + DEP["E"] + ")", True, "再発行後の請求番号へ付け替える"),
    ("入金明細の入力漏れ", "=COUNTIF(" + DEP["I"] + ',"要入力")', None, True, "入金日・入金額を入力"),
    ("section", "■ 請求漏れ(対象月)"),
    ("毎月請求先で対象月の請求が未登録", "=COUNTIF(" + rng(S_BAL, "J") + ',"請求漏れ候補")', None, True,
     "「取引先別残高」の請求漏れ候補を確認し、請求するか理由を備考に記録"),
    ("マスタの取引先コード重複", "=COUNTIF(" + CLI["L"] + ',"コード重複")', None, True, "取引先マスタを修正"),
    ("section", "■ 月次の数字(対象月)"),
    ("対象月の請求(取消除く)", "=COUNTIFS(" + INV["B"] + ',">="&' + TARGET_MONTH + "," + INV["B"] + ',"<="&EOMONTH(' + TARGET_MONTH
     + ',0),' + INV["T"] + ',"<>取消")',
     "=SUMIFS(" + INV["I"] + "," + INV["B"] + ',">="&' + TARGET_MONTH + "," + INV["B"] + ',"<="&EOMONTH('
     + TARGET_MONTH + ',0),' + INV["T"] + ',"<>取消")', None,
     "契約・受注・納品の一覧と件数を突合し、請求漏れがないか確認"),
    ("対象月の入金", "=SUMPRODUCT((" + DEP["B"] + ">=" + TARGET_MONTH + ")*(" + DEP["B"] + "<=EOMONTH("
     + TARGET_MONTH + ",0))*(" + DEP["E"] + "<>\"\"))",
     "=SUMIFS(" + DEP["E"] + "," + DEP["B"] + ',">="&' + TARGET_MONTH + "," + DEP["B"] + ',"<="&EOMONTH('
     + TARGET_MONTH + ",0))", None, "通帳の対象月入金合計(請求入金分)と一致するか確認"),
    ("未回収残高 合計(基準日時点)", "=COUNTIFS(" + INV["S"] + ',">0",' + INV["T"] + ',"<>取消")+COUNTIFS(' + INV["S"] + ',"<0",' + INV["T"] + ',"<>取消")',
     "=SUMIF(" + INV["T"] + ',"<>取消",' + INV["S"] + ")", None, "下の会計ソフトの売掛金残高と照合"),
]
r = 8
row_of = {}
for item in checks:
    if item[0] == "section":
        c = ck.cell(row=r, column=2, value=item[1])
        c.font = F_BOLD
        for col in range(2, 7):
            ck.cell(row=r, column=col).fill = FILL_SECTION
        r += 1
        continue
    label, count_f, amount_f, judged, action = item
    row_of[label] = r
    ck.cell(row=r, column=2, value=label)
    ck.cell(row=r, column=3, value=count_f).number_format = '#,##0"件"'
    if amount_f:
        ck.cell(row=r, column=4, value=amount_f).number_format = FMT_YEN
    if judged is True:
        ck.cell(row=r, column=5, value=f'=IF(C{r}=0,"OK","要対応")')
    elif judged is False:
        ck.cell(row=r, column=5, value="参考")
    ck.cell(row=r, column=6, value=action)
    for col in range(2, 7):
        c = ck.cell(row=r, column=col)
        c.font, c.border = F_BASE, BORDER
        if col == 6:
            c.alignment = WRAP
    r += 1

# 会計ソフトとの照合
r += 1
ck.cell(row=r, column=2, value="■ 会計ソフトとの残高照合(月次)").font = F_BOLD
for col in range(2, 7):
    ck.cell(row=r, column=col).fill = FILL_SECTION
r += 1
ck.cell(row=r, column=2, value="会計ソフトの売掛金残高(基準日時点)")
acct = ck.cell(row=r, column=4)
acct.fill, acct.number_format = FILL_INPUT, FMT_YEN
acct.protection = Protection(locked=False)
ck.cell(row=r, column=6, value="会計ソフトの売掛金残高を入力(基準日と同じ日付時点)")
acct_row = r
r += 1
ck.cell(row=r, column=2, value="差額(台帳 − 会計ソフト)")
ck.cell(row=r, column=4, value=f'=IF(D{acct_row}="","",D{row_of["未回収残高 合計(基準日時点)"]}-D{acct_row})'
        ).number_format = FMT_YEN
ck.cell(row=r, column=5, value=f'=IF(D{acct_row}="","未入力",IF(D{r}=0,"OK","要対応"))')
ck.cell(row=r, column=6, value="差額があれば、台帳への登録漏れ・会計側の計上漏れ・消込漏れのいずれかを調査")
for rr in (acct_row, r):
    for col in range(2, 7):
        c = ck.cell(row=rr, column=col)
        c.font, c.border = F_BASE, BORDER
ck.conditional_formatting.add(f"E8:E{r}", FormulaRule(
    formula=['$E8="要対応"'], fill=PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD"),
    font=Font(name=FONT, bold=True, color="9C0006")))
ck.conditional_formatting.add(f"E8:E{r}", FormulaRule(
    formula=['$E8="OK"'], fill=PatternFill("solid", fgColor="E2EFDA", bgColor="E2EFDA"),
    font=Font(name=FONT, bold=True, color="375623")))
ck.conditional_formatting.add(f"E8:E{r}", FormulaRule(
    formula=['$E8="未入力"'], fill=PatternFill("solid", fgColor="FFFF00", bgColor="FFFF00"),
    font=Font(name=FONT, bold=True, color="9C0006")))
r += 2
ck.cell(row=r, column=2, value="総合判定").font = F_BOLD
ck.cell(row=r, column=3, value=f'=IF(COUNTIF(E8:E{r - 2},"要対応")+COUNTIF(E8:E{r - 2},"未入力")=0,'
        f'"すべてOK","要対応あり")').font = F_BOLD
ck.cell(row=r, column=3).border = BORDER
ck.conditional_formatting.add(f"C{r}", FormulaRule(
    formula=[f'$C${r}="要対応あり"'], fill=PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD"),
    font=Font(name=FONT, bold=True, color="9C0006")))
ck.conditional_formatting.add(f"C{r}", FormulaRule(
    formula=[f'$C${r}="すべてOK"'], fill=PatternFill("solid", fgColor="E2EFDA", bgColor="E2EFDA"),
    font=Font(name=FONT, bold=True, color="375623")))
protect(ck)

# ---------------------------------------------------------------- 取引先別残高
bl = wb.create_sheet("取引先別残高")
title(bl, "取引先別残高",
      "取引先マスタの順に自動表示(入力不要)。超過日数別の残高と、毎月請求先の対象月請求漏れ候補を表示します。")


bal_cols = [
    ("取引先コード", 11, "auto", None, '=IF(' + S_CLI + '!A{r}="","",' + S_CLI + '!A{r})', None),
    ("取引先名", 26, "auto", None, '=IF(A{r}="","",' + S_CLI + '!B{r})', None),
    ("営業担当", 10, "auto", None, '=IF(A{r}="","",' + S_CLI + '!J{r}&"")', None),
    ("未回収残高", 13, "auto", FMT_YEN,
     '=IF(A{r}="","",SUMIFS(' + INV["S"] + ',' + INV["C"] + ',A{r},' + INV["T"] + ',"<>取消"))', None),
    ("期日内", 12, "auto", FMT_YEN, '=IF(A{r}="","",D{r}-SUM(F{r}:H{r}))', None),
    ("超過 1〜30日", 12, "auto", FMT_YEN,
     '=IF(A{r}="","",SUMIFS(' + INV["S"] + ',' + INV["C"] + ',A{r},' + INV["V"] + ',">=1",' + INV["V"] + ',"<=30"))',
     None),
    ("超過 31〜60日", 12, "auto", FMT_YEN,
     '=IF(A{r}="","",SUMIFS(' + INV["S"] + ',' + INV["C"] + ',A{r},' + INV["V"] + ',">=31",' + INV["V"] + ',"<=60"))',
     None),
    ("超過 61日以上", 12, "auto", FMT_YEN,
     '=IF(A{r}="","",SUMIFS(' + INV["S"] + ',' + INV["C"] + ',A{r},' + INV["V"] + ',">=61"))', None),
    ("対象月の請求件数", 11, "auto", '0"件"',
     '=IF(A{r}="","",COUNTIFS(' + INV["C"] + ',A{r},' + INV["B"] + ',">="&' + TARGET_MONTH + ','
     + INV["B"] + ',"<="&EOMONTH(' + TARGET_MONTH + ',0),' + INV["T"] + ',"<>取消"))',
     "「チェック」シートの対象月に請求日がある請求の件数(取消除く)"),
    ("請求漏れ判定", 13, "auto", None,
     '=IF(A{r}="","",IF(AND(' + S_CLI + '!G{r}="有",I{r}=0),"請求漏れ候補",""))',
     "取引先マスタで毎月請求=有 なのに対象月の請求がない"),
]
table(bl, bal_cols, FIRST, CLI_LAST)
bl.auto_filter.ref = None
tot = CLI_LAST + 1
bl.cell(row=tot, column=2, value="合計").font = F_BOLD
for col in "DEFGH":
    c = bl[f"{col}{tot}"]
    c.value = f"=SUM({col}{FIRST}:{col}{CLI_LAST})"
    c.font, c.number_format, c.border = F_BOLD, FMT_YEN, BORDER
# 合計を上部にも表示
bl["C3"] = "合計"
bl["C3"].font = F_BOLD
for col in "DEFGH":
    c = bl[f"{col}3"]
    c.value = f"={col}{tot}"
    c.font, c.number_format, c.border = F_BOLD, FMT_YEN, BORDER
bl.conditional_formatting.add(f"J{FIRST}:J{CLI_LAST}", FormulaRule(
    formula=[f'$J{FIRST}="請求漏れ候補"'], fill=PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD"),
    font=Font(name=FONT, bold=True, color="9C0006")))
for col in "FGH":
    bl.conditional_formatting.add(f"{col}{FIRST}:{col}{CLI_LAST}", FormulaRule(
        formula=[f'N(${col}{FIRST})>0'], font=Font(name=FONT, bold=True, color="C00000")))
protect(bl)

# ---------------------------------------------------------------- 月次チェック記録
mr = wb.create_sheet("月次チェック記録")
title(mr, "月次チェック記録",
      "毎月、締め後(例: 翌月5営業日以内)に実施。各項目を確認したら「済」を選択し、実施者・承認者を記入します。")
items = [
    ("対象月", 11, FMT_MONTH, None),
    ("①請求漏れ突合\n(受注・納品一覧と台帳)", 16, None, "契約・受注・納品・作業報告の一覧と、対象月の請求台帳を1件ずつ突合"),
    ("②未送付 0件", 10, None, None),
    ("③要入力 0件", 10, None, None),
    ("④未消込入金 0件\n(番号誤り含む)", 13, None, None),
    ("⑤期日超過の督促・\n状況確認", 14, None, "期日超過の全件について、担当者が先方に確認し備考に結果を記録"),
    ("⑥会計ソフト残高\nとの照合(差額0)", 14, None, None),
    ("⑦毎月請求先の\n請求漏れ候補 0件", 14, None, None),
    ("実施者", 10, None, None),
    ("実施日", 11, FMT_DATE, None),
    ("承認者(責任者)", 12, None, "実施者とは別の人が、チェックシートの画面を見て承認"),
    ("承認日", 11, FMT_DATE, None),
    ("期日超過の件数・金額\n/ 特記事項", 40, None, None),
]
for i, (head, width, fmt, comment) in enumerate(items, start=1):
    col = get_column_letter(i)
    c = mr.cell(row=HDR, column=i, value=head)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
    if comment:
        c.comment = Comment(comment, "管理表")
    mr.column_dimensions[col].width = width
    for rr in range(FIRST, FIRST + 36):
        cell = mr.cell(row=rr, column=i)
        cell.font, cell.fill, cell.border = F_BASE, FILL_INPUT, BORDER
        cell.protection = Protection(locked=False)
        if fmt:
            cell.number_format = fmt
        if 2 <= i <= 8:
            cell.alignment = CENTER
mr.row_dimensions[HDR].height = 46
mr.freeze_panes = "B5"
for k in range(36):
    y, m = divmod(9 - 1 + k, 12)
    mr.cell(row=FIRST + k, column=1, value=date(2026 + y, m + 1, 1))
add_list(mr, f"B{FIRST}:H{FIRST + 35}", list_src("K"))
add_date(mr, f"J{FIRST}:J{FIRST + 35}")
add_date(mr, f"L{FIRST}:L{FIRST + 35}")
mr.conditional_formatting.add(f"A{FIRST}:M{FIRST + 35}", FormulaRule(
    formula=[f'AND($A{FIRST}<>"",$A{FIRST}<' + S_DASH + '!$C$4,$L{FIRST}="")'],
    fill=PatternFill("solid", fgColor="FCE4D6", bgColor="FCE4D6")))
mr["N4"] = "※ 過去月で承認日が空欄の行はオレンジ表示"
mr["N4"].font = F_NOTE
protect(mr)

# シート順
order = ["はじめに", "チェック", "請求台帳", "入金明細", "取引先別残高", "取引先マスタ", "月次チェック記録", "設定"]
wb._sheets = [wb[n] for n in order]
for s in wb.worksheets:
    s.sheet_properties.tabColor = {
        "はじめに": "7F7F7F", "チェック": "C00000", "請求台帳": "1F4E78", "入金明細": "1F4E78",
        "取引先別残高": "548235", "取引先マスタ": "BF8F00", "月次チェック記録": "C00000", "設定": "7F7F7F",
    }[s.title]

wb.save(OUT)
print("saved", OUT)
