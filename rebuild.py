#!/usr/bin/env python3
"""家計簿 全シート再構築（5月開始・数式修正・目標/実績/残金）"""

import json
from pathlib import Path
from datetime import date
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
CREDENTIALS_PATH = SCRIPT_DIR / "credentials.json"
TOKEN_PATH = SCRIPT_DIR / "token.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

YEAR = date.today().year
CUR_MONTH = date.today().month  # 5

# 5月〜12月のみ表示
START_MONTH = 5
MONTHS = [f"{m}月" for m in range(START_MONTH, 13)]   # ["5月"..."12月"]
MONTH_NUMS = list(range(START_MONTH, 13))              # [5,6,...,12]
NUM_MONTHS = len(MONTHS)                               # 8

CATEGORIES = [
    "家賃","光熱費","通信費","保育料","保険","外貨積立","税金",
    "医療費","衣料美容費","食費","日用品","健太郎お小遣い",
    "藍子お小遣い","陽和お小遣い","サブスク代","外食費","自動車","ペット","その他",
]

CATEGORY_COLORS = {
    "家賃":"#FADBD8","光熱費":"#FDEBD0","通信費":"#D6EAF8",
    "保育料":"#FDEDEC","保険":"#E8DAEF","外貨積立":"#D1F2EB",
    "税金":"#EAECEE","医療費":"#F9EBEA","衣料美容費":"#FDEBF7",
    "食費":"#D5F5E3","日用品":"#EAFAF1","健太郎お小遣い":"#DAE8FC",
    "藍子お小遣い":"#E8D5FC","陽和お小遣い":"#FEF9E7",
    "サブスク代":"#D5D8DC","外食費":"#FDEBD0","自動車":"#D6DBDF",
    "ペット":"#D4EFDF","その他":"#F2F3F4",
}

DEFAULT_BUDGETS = {
    "家賃":85600,      # 住宅ローン実績
    "光熱費":15000,
    "通信費":20000,    # 見直し後目標額
    "保育料":28200,    # 固定
    "保険":24000,      # メットライフ¥5,000+ジブラルタ~¥18,500（損保ジャパンは自動車へ）
    "外貨積立":20000,
    "税金":91000,      # 国民年金¥16,980+国保¥35,700+住民税¥23,600+所得税¥15,000（仮算出）
    "医療費":5000,
    "衣料美容費":20000,
    "食費":30000,
    "日用品":7000,
    "健太郎お小遣い":10000,
    "藍子お小遣い":10000,
    "陽和お小遣い":10000,
    "サブスク代":34000, # GWS¥2,500+エニタイム¥16,000+Claude¥3,000+Miro¥1,250+Adobe¥6,600+Canva¥1,500+SONY¥550+Netflix¥890+AppleMusic¥1,680
    "外食費":15000,
    "自動車":56000,    # オートローン¥27,338+ガソリン¥15,000+損保ジャパン¥13,283
    "ペット":8900,     # クオーク犬ローン（〜8月）
    "その他":20000,
}

WIFE_INCOME    = 250000
HUSBAND_INCOME = 250000
TOTAL_INCOME   = WIFE_INCOME + HUSBAND_INCOME

NAVY  = "#1E3A5F"
BLUE  = "#2E86C1"
BLUE2 = "#1A5276"
GREEN = "#1A5E39"
W     = {"red":1,"green":1,"blue":1}


def hc(h):
    h = h.lstrip("#")
    return {"red":int(h[0:2],16)/255,"green":int(h[2:4],16)/255,"blue":int(h[4:6],16)/255}

def col(i): return chr(ord('A') + i)  # 0→A, 1→B ...

def upd(ws, range_name, values):
    """数式を確実に評価させるための update ラッパー"""
    ws.spreadsheet.values_update(
        f"'{ws.title}'!{range_name}",
        params={"valueInputOption": "USER_ENTERED"},
        body={"values": values}
    )

def get_credentials():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.write_text(creds.to_json())
    return creds


# ══════════════════════════════════════════════════════
# 明細シート
# ══════════════════════════════════════════════════════
def setup_meisai(spreadsheet):
    try:
        ws = spreadsheet.worksheet("明細")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="明細", rows=2000, cols=7)

    upd(ws, "A1:G1", [["日付","曜日","店舗名","品目","費目","金額(税込)","メモ"]])

    reqs = [
        # ヘッダー書式
        {"repeatCell":{
            "range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":7},
            "cell":{"userEnteredFormat":{
                "backgroundColor":hc(NAVY),
                "textFormat":{"foregroundColor":W,"bold":True,"fontSize":11},
                "horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"
            }},
            "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":40},"fields":"pixelSize"}},
        {"updateSheetProperties":{"properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":1}},"fields":"gridProperties.frozenRowCount"}},
    ]
    for i, px in enumerate([100,50,150,150,130,110,220]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":px},"fields":"pixelSize"}})
    # 金額書式
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":2000,"startColumnIndex":5,"endColumnIndex":6},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}})
    # 費目別カラー
    for idx, cat in enumerate(CATEGORIES):
        reqs.append({"addConditionalFormatRule":{
            "rule":{
                "ranges":[{"sheetId":ws.id,"startRowIndex":1,"startColumnIndex":0,"endColumnIndex":7}],
                "booleanRule":{
                    "condition":{"type":"CUSTOM_FORMULA","values":[{"userEnteredValue":f'=$E2="{cat}"'}]},
                    "format":{"backgroundColor":hc(CATEGORY_COLORS[cat])}
                }
            },"index":idx
        }})
    spreadsheet.batch_update({"requests":reqs})
    print("  ✓ 明細")
    return ws


# ══════════════════════════════════════════════════════
# 月別集計シート（5月〜12月）
# ══════════════════════════════════════════════════════
def setup_monthly(spreadsheet):
    try:
        ws = spreadsheet.worksheet("月別集計")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="月別集計", rows=30, cols=12)

    header = ["費目"] + MONTHS + ["合計(5-12月)"]
    upd(ws, f"A1:{col(NUM_MONTHS+1)}1", [header])

    rows = []
    for cat in CATEGORIES:
        row = [cat]
        for m in MONTH_NUMS:
            row.append(f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")*(MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)')
        row.append(f'=SUM(B{len(rows)+2}:{col(NUM_MONTHS)}{len(rows)+2})')
        rows.append(row)

    total = ["合計"]
    for i in range(1, NUM_MONTHS+1):
        total.append(f'=SUM({col(i)}2:{col(i)}{len(CATEGORIES)+1})')
    total.append(f'=SUM({col(NUM_MONTHS+1)}2:{col(NUM_MONTHS+1)}{len(CATEGORIES)+1})')
    rows.append(total)

    upd(ws, f"A2:{col(NUM_MONTHS+1)}{len(CATEGORIES)+2}", rows)

    n = len(CATEGORIES)
    reqs = [
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":10},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":n+1,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True},"backgroundColor":hc("#EBF5FB")}},"fields":"userEnteredFormat(textFormat,backgroundColor)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":n+1,"endRowIndex":n+2,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE),"textFormat":{"foregroundColor":W,"bold":True},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":n+2,"startColumnIndex":1,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":n+2},"properties":{"pixelSize":34},"fields":"pixelSize"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":40},"fields":"pixelSize"}},
        {"updateSheetProperties":{"properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":1,"frozenColumnCount":1}},"fields":"gridProperties(frozenRowCount,frozenColumnCount)"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":0,"endIndex":1},"properties":{"pixelSize":165},"fields":"pixelSize"}},
    ]
    for i in range(1, NUM_MONTHS+2):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":88},"fields":"pixelSize"}})
    for i in range(n):
        reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":i+1,"endRowIndex":i+2,"startColumnIndex":1,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#F4F9FD" if i%2==0 else "#FFFFFF")}},"fields":"userEnteredFormat(backgroundColor)"}})
    reqs.append({"updateBorders":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":n+2,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"innerHorizontal":{"style":"SOLID","color":hc("#D5D8DC")},"innerVertical":{"style":"SOLID","color":hc("#D5D8DC")},"top":{"style":"SOLID_MEDIUM","color":hc(BLUE)},"bottom":{"style":"SOLID_MEDIUM","color":hc(BLUE)},"left":{"style":"SOLID_MEDIUM","color":hc(BLUE)},"right":{"style":"SOLID_MEDIUM","color":hc(BLUE)}}})

    spreadsheet.batch_update({"requests":reqs})
    print("  ✓ 月別集計")
    return ws


# ══════════════════════════════════════════════════════
# 予算プランシート（目標 / 実績 / 残金）
# ══════════════════════════════════════════════════════
def setup_budget_plan(spreadsheet):
    try:
        ws = spreadsheet.worksheet("予算プラン")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="予算プラン", rows=60, cols=10)

    n = len(CATEGORIES)

    # ── タイトル・基本設定 ──
    upd(ws, "A1", [[f"{YEAR}年 家計予算プラン（5月〜12月）"]])
    upd(ws, "A3:B3", [["基本設定",""]])
    upd(ws, "A4:B7", [
        ["妻（藍子）月収",    WIFE_INCOME],
        ["夫（健太郎）月収",  HUSBAND_INCOME],
        ["月収合計",          "=B4+B5"],
        ["年収合計（8か月）", "=B6*8"],
    ])

    # ── 費目別予算テーブル（5月〜12月 各月：目標/実績/残金） ──
    # 列構成: A=費目, B=月次目標, C〜J = 5月〜12月（各月に実績・残金の2列）
    # でも列が多くなりすぎるので、シンプルに「今月（5月）」のみ詳細表示 + 月別実績横並び

    # ── シンプル構成: 費目 | 月次目標 | 今月実績(5月) | 残金 | 達成率 ──
    upd(ws, "A9", [["費目別 目標・実績・残金"]])
    upd(ws, "A10:E10", [["費目","月次目標","今月実績","残金(目標-実績)","達成率"]])

    cat_rows = []
    for i, cat in enumerate(CATEGORIES):
        r = 11 + i
        budget = DEFAULT_BUDGETS.get(cat, 0)
        actual = (f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")'
                  f'*(MONTH(明細!$A$2:$A$2000)={CUR_MONTH})'
                  f'*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)')
        cat_rows.append([
            cat,
            budget,
            actual,
            f"=B{r}-C{r}",
            f'=IFERROR(C{r}/B{r},0)',
        ])
    upd(ws, f"A11:E{10+n}", cat_rows)

    # 合計行
    sr = 11 + n
    upd(ws, f"A{sr}:E{sr}", [[
        "合計",
        f"=SUM(B11:B{10+n})",
        f"=SUM(C11:C{10+n})",
        f"=SUM(D11:D{10+n})",
        f'=IFERROR(C{sr}/B{sr},0)',
    ]])

    # 貯蓄分析
    s = sr + 2
    upd(ws, f"A{s}", [["貯蓄分析"]])
    upd(ws, f"A{s+1}:B{s+5}", [
        ["月収合計",         "=$B$6"],
        ["月次支出予算合計", f"=B{sr}"],
        ["月次貯蓄可能額",   f"=$B$6-B{sr}"],
        ["貯蓄率",           f"=IFERROR(B{s+3}/$B$6,0)"],
        ["年間貯蓄予測(8か月)", f"=B{s+3}*8"],
    ])

    # ── 月別実績サマリー（5月〜12月を横並び） ──
    ms = s + 8
    upd(ws, f"A{ms}", [["月別支出サマリー（5月〜12月）"]])
    upd(ws, f"A{ms+1}:{col(NUM_MONTHS+1)}{ms+1}", [["費目"] + MONTHS])
    month_rows = []
    for cat in CATEGORIES:
        row = [cat]
        for m in MONTH_NUMS:
            row.append(f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")*(MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)')
        month_rows.append(row)
    # 月合計行
    m_total = ["月合計"]
    for j in range(NUM_MONTHS):
        m_total.append(f'=SUM({col(j+1)}{ms+3}:{col(j+1)}{ms+2+n})')
    month_rows.append(m_total)
    upd(ws, f"A{ms+2}:{col(NUM_MONTHS)}{ms+2+n}", month_rows)

    # ── 書式 ──
    reqs = []

    # タイトル
    reqs += [
        {"unmergeCells":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":60,"startColumnIndex":0,"endColumnIndex":10}}},
        {"mergeCells":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":10},"mergeType":"MERGE_ALL"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":10},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":16},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":60},"fields":"pixelSize"}},
    ]

    # 基本設定見出し
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":2,"endRowIndex":3,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":11}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":3,"endRowIndex":7,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#EBF5FB")}},"fields":"userEnteredFormat(backgroundColor)"}})
    for r,bg in [(3,"#FDEBF7"),(4,"#D6EAF8")]:
        reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc(bg),"textFormat":{"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})

    # 費目別テーブル見出し
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":8,"endRowIndex":9,"startColumnIndex":0,"endColumnIndex":10},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
    # テーブルヘッダー
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":10,"startColumnIndex":0,"endColumnIndex":5},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":10},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}})
    # 交互行
    for i in range(n):
        reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":10+i,"endRowIndex":11+i,"startColumnIndex":0,"endColumnIndex":5},"cell":{"userEnteredFormat":{"backgroundColor":hc("#F4F9FD" if i%2==0 else "#FFFFFF")}},"fields":"userEnteredFormat(backgroundColor)"}})
    # 合計行
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":10+n,"endRowIndex":11+n,"startColumnIndex":0,"endColumnIndex":5},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":11}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})

    # 残金列: 負=赤背景, 正=緑背景
    reqs.append({"addConditionalFormatRule":{"rule":{"ranges":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+n,"startColumnIndex":3,"endColumnIndex":4}],"booleanRule":{"condition":{"type":"NUMBER_LESS","values":[{"userEnteredValue":"0"}]},"format":{"backgroundColor":hc("#FADBD8"),"textFormat":{"foregroundColor":hc("#C0392B"),"bold":True}}}},"index":0}})
    reqs.append({"addConditionalFormatRule":{"rule":{"ranges":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+n,"startColumnIndex":3,"endColumnIndex":4}],"booleanRule":{"condition":{"type":"NUMBER_GREATER_THAN_EQ","values":[{"userEnteredValue":"0"}]},"format":{"backgroundColor":hc("#EAFAF1"),"textFormat":{"foregroundColor":hc("#1A7431"),"bold":True}}}},"index":1}})
    # 達成率 100%超 = 赤
    reqs.append({"addConditionalFormatRule":{"rule":{"ranges":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+n,"startColumnIndex":4,"endColumnIndex":5}],"booleanRule":{"condition":{"type":"NUMBER_GREATER","values":[{"userEnteredValue":"1"}]},"format":{"backgroundColor":hc("#FADBD8")}}},"index":2}})

    # 貯蓄セクション見出し
    s0 = sr + 1  # 0-indexed
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":s0,"endRowIndex":s0+1,"startColumnIndex":0,"endColumnIndex":10},"cell":{"userEnteredFormat":{"backgroundColor":hc(GREEN),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":s0+1,"endRowIndex":s0+6,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#EAFAF1")}},"fields":"userEnteredFormat(backgroundColor)"}})

    # 月別サマリー見出し
    ms0 = ms - 1  # 0-indexed
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":ms0,"endRowIndex":ms0+1,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+1},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":12}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":ms0+1,"endRowIndex":ms0+2,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+1},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":ms0+n+2,"endRowIndex":ms0+n+3,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+1},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE),"textFormat":{"foregroundColor":W,"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}})

    # 数値書式 ¥
    money_cells = [
        (3,8,1,2), (9,10+n+1,1,4), (s0+1,s0+4,1,2), (s0+5,s0+6,1,2),
        (ms0+2,ms0+2+n+1,1,NUM_MONTHS+1),
    ]
    for r0,r1,c0,c1 in money_cells:
        reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":r0,"endRowIndex":r1,"startColumnIndex":c0,"endColumnIndex":c1},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}})
    # % 書式
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":10+n+1,"startColumnIndex":4,"endColumnIndex":5},"cell":{"userEnteredFormat":{"numberFormat":{"type":"PERCENT","pattern":"0%"},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}})
    reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":s0+4,"endRowIndex":s0+5,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"numberFormat":{"type":"PERCENT","pattern":"0.0%"}}},"fields":"userEnteredFormat(numberFormat)"}})

    # 列幅
    for ci,px in enumerate([170,110,110,120,80]):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":ci,"endIndex":ci+1},"properties":{"pixelSize":px},"fields":"pixelSize"}})

    # 固定（費目テーブルヘッダー行まで）
    reqs.append({"updateSheetProperties":{"properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":10}},"fields":"gridProperties.frozenRowCount"}})

    # ボーダー
    reqs.append({"updateBorders":{"range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":10+n+1,"startColumnIndex":0,"endColumnIndex":5},"innerHorizontal":{"style":"SOLID","color":hc("#D5D8DC")},"innerVertical":{"style":"SOLID","color":hc("#D5D8DC")},"top":{"style":"SOLID_MEDIUM","color":hc(BLUE)},"bottom":{"style":"SOLID_MEDIUM","color":hc(BLUE)},"left":{"style":"SOLID_MEDIUM","color":hc(BLUE)},"right":{"style":"SOLID_MEDIUM","color":hc(BLUE)}}})

    spreadsheet.batch_update({"requests":reqs})
    print("  ✓ 予算プラン")
    return ws


# ══════════════════════════════════════════════════════
# 収支管理シート（5月〜12月）
# ══════════════════════════════════════════════════════
def setup_shuushi(spreadsheet):
    try:
        ws = spreadsheet.worksheet("収支管理")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="収支管理", rows=20, cols=12)

    end_col = col(NUM_MONTHS+1)  # J (for 8 months + label + total)

    upd(ws, "A1", [[f"{YEAR}年 月次収支管理（5月〜12月）"]])
    upd(ws, f"A2:{end_col}2", [["項目"] + MONTHS + ["合計"]])

    # 妻収入
    wife = ["妻（藍子）収入"] + [WIFE_INCOME]*NUM_MONTHS + [f"=SUM(B3:{col(NUM_MONTHS)}3)"]
    upd(ws, f"A3:{end_col}3", [wife])

    # 夫収入
    husb = ["夫（健太郎）収入"] + [HUSBAND_INCOME]*NUM_MONTHS + [f"=SUM(B4:{col(NUM_MONTHS)}4)"]
    upd(ws, f"A4:{end_col}4", [husb])

    # 収入合計
    inc_total = ["収入合計"]
    for i in range(NUM_MONTHS):
        c = col(i+1)
        inc_total.append(f"={c}3+{c}4")
    inc_total.append(f"=SUM(B5:{col(NUM_MONTHS)}5)")
    upd(ws, f"A5:{end_col}5", [inc_total])

    # 支出合計（明細から）
    expense = ["支出合計（実績）"]
    for m in MONTH_NUMS:
        expense.append(f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)')
    expense.append(f"=SUM(B6:{col(NUM_MONTHS)}6)")
    upd(ws, f"A6:{end_col}6", [expense])

    # 収支差額
    diff = ["収支差額"]
    for i in range(NUM_MONTHS):
        c = col(i+1)
        diff.append(f"={c}5-{c}6")
    diff.append(f"=SUM(B7:{col(NUM_MONTHS)}7)")
    upd(ws, f"A7:{end_col}7", [diff])

    # 累計
    cum = ["累計収支", "=B7"]
    for i in range(1, NUM_MONTHS):
        cum.append(f"={col(i)}8+{col(i+1)}7")
    cum.append("")
    upd(ws, f"A8:{end_col}8", [cum])

    reqs = [
        {"unmergeCells":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":20,"startColumnIndex":0,"endColumnIndex":15}}},
        {"mergeCells":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"mergeType":"MERGE_ALL"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":16},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":56},"fields":"pixelSize"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":2,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":10},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":2,"endRowIndex":3,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#FDEBF7"),"textFormat":{"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":3,"endRowIndex":4,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#D6EAF8"),"textFormat":{"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":4,"endRowIndex":5,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#D5F5E3"),"textFormat":{"bold":True,"fontSize":11}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":5,"endRowIndex":6,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#FDEBD0"),"textFormat":{"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":6,"endRowIndex":7,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":11}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":7,"endRowIndex":8,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        # 数値書式
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":2,"endRowIndex":9,"startColumnIndex":1,"endColumnIndex":NUM_MONTHS+2},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}},
        # 収支差額マイナス = 赤
        {"addConditionalFormatRule":{"rule":{"ranges":[{"sheetId":ws.id,"startRowIndex":6,"endRowIndex":7,"startColumnIndex":1,"endColumnIndex":NUM_MONTHS+1}],"booleanRule":{"condition":{"type":"NUMBER_LESS","values":[{"userEnteredValue":"0"}]},"format":{"textFormat":{"foregroundColor":hc("#E74C3C"),"bold":True}}}},"index":0}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":1,"endIndex":9},"properties":{"pixelSize":36},"fields":"pixelSize"}},
        {"updateSheetProperties":{"properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":2}},"fields":"gridProperties.frozenRowCount"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":0,"endIndex":1},"properties":{"pixelSize":175},"fields":"pixelSize"}},
    ]
    for i in range(1, NUM_MONTHS+2):
        reqs.append({"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":92},"fields":"pixelSize"}})
    reqs.append({"updateBorders":{"range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":9,"startColumnIndex":0,"endColumnIndex":NUM_MONTHS+2},"innerHorizontal":{"style":"SOLID","color":hc("#D5D8DC")},"innerVertical":{"style":"SOLID","color":hc("#D5D8DC")}}})

    spreadsheet.batch_update({"requests":reqs})
    print("  ✓ 収支管理")
    return ws


# ══════════════════════════════════════════════════════
# ダッシュボード
# ══════════════════════════════════════════════════════
def setup_dashboard(spreadsheet):
    try:
        ws = spreadsheet.worksheet("ダッシュボード")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="ダッシュボード", rows=60, cols=12)

    n = len(CATEGORIES)
    m = CUR_MONTH
    # 収支管理の該当月列 (5月=B列=index1)
    income_col = col(m - START_MONTH + 1)

    upd(ws, "A1", [[f"{YEAR}年 家計簿ダッシュボード"]])

    # サマリーカード
    upd(ws, "A3:B8", [
        ["今月", f"{m}月"],
        ["💸 今月の支出",
            f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)'],
        ["💰 今月の収入",   f"=収支管理!{income_col}5"],
        ["📊 収支差額",     "=B5-B4"],
        ["📦 今月の件数",
            f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={YEAR})*(明細!$A$2:$A$2000<>""))'],
        ["📈 今月の平均/件", "=IFERROR(B4/B6,0)"],
    ])

    # 費目別（チャート用）
    upd(ws, "A10:B10", [["費目", f"{m}月支出"]])
    cat_rows = []
    for cat in CATEGORIES:
        cat_rows.append([cat,
            f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")*(MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)'])
    upd(ws, f"A11:B{10+n}", cat_rows)

    # 月別推移（チャート用）
    upd(ws, "D10:E10", [["月", "支出合計"]])
    month_rows = []
    for mo in MONTH_NUMS:
        month_rows.append([f"{mo}月",
            f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={mo})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)'])
    upd(ws, f"D11:E{10+NUM_MONTHS}", month_rows)

    reqs = [
        # タイトル
        {"unmergeCells":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":60,"startColumnIndex":0,"endColumnIndex":12}}},
        {"mergeCells":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":12},"mergeType":"MERGE_ALL"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":12},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":18},"horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},"properties":{"pixelSize":64},"fields":"pixelSize"}},
        # 今月ヘッダー
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":2,"endRowIndex":3,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE),"textFormat":{"foregroundColor":W,"bold":True,"fontSize":11},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}},
        # 支出 赤
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":3,"endRowIndex":4,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#FADBD8"),"textFormat":{"bold":True,"fontSize":14,"foregroundColor":hc("#C0392B")}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        # 収入 緑
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":4,"endRowIndex":5,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#EAFAF1"),"textFormat":{"bold":True,"fontSize":14,"foregroundColor":hc("#1A7431")}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        # 差額 紺
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":5,"endRowIndex":6,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc(NAVY),"textFormat":{"bold":True,"fontSize":14,"foregroundColor":W}}},"fields":"userEnteredFormat(backgroundColor,textFormat)"}},
        # 件数・平均
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":6,"endRowIndex":8,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#EBF5FB")}},"fields":"userEnteredFormat(backgroundColor)"}},
        # 数値書式
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":3,"endRowIndex":6,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"}}},"fields":"userEnteredFormat(numberFormat)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":7,"endRowIndex":8,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"}}},"fields":"userEnteredFormat(numberFormat)"}},
        # 差額マイナス = 赤
        {"addConditionalFormatRule":{"rule":{"ranges":[{"sheetId":ws.id,"startRowIndex":5,"endRowIndex":6,"startColumnIndex":1,"endColumnIndex":2}],"booleanRule":{"condition":{"type":"NUMBER_LESS","values":[{"userEnteredValue":"0"}]},"format":{"backgroundColor":hc("#FADBD8"),"textFormat":{"foregroundColor":hc("#C0392B"),"bold":True}}}},"index":0}},
        # 費目テーブル
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":10,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+n,"startColumnIndex":1,"endColumnIndex":2},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}},
        # 月次テーブル
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":10,"startColumnIndex":3,"endColumnIndex":5},"cell":{"userEnteredFormat":{"backgroundColor":hc(BLUE2),"textFormat":{"foregroundColor":W,"bold":True},"horizontalAlignment":"CENTER"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"}},
        {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+NUM_MONTHS,"startColumnIndex":4,"endColumnIndex":5},"cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},"horizontalAlignment":"RIGHT"}},"fields":"userEnteredFormat(numberFormat,horizontalAlignment)"}},
        # 列幅
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":0,"endIndex":1},"properties":{"pixelSize":175},"fields":"pixelSize"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":1,"endIndex":2},"properties":{"pixelSize":120},"fields":"pixelSize"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":2,"endIndex":3},"properties":{"pixelSize":20},"fields":"pixelSize"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":3,"endIndex":4},"properties":{"pixelSize":70},"fields":"pixelSize"}},
        {"updateDimensionProperties":{"range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":4,"endIndex":5},"properties":{"pixelSize":110},"fields":"pixelSize"}},
    ]
    for i in range(n):
        reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":10+i,"endRowIndex":11+i,"startColumnIndex":0,"endColumnIndex":2},"cell":{"userEnteredFormat":{"backgroundColor":hc("#F4F9FD" if i%2==0 else "#FFFFFF")}},"fields":"userEnteredFormat(backgroundColor)"}})
    for i in range(NUM_MONTHS):
        reqs.append({"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":10+i,"endRowIndex":11+i,"startColumnIndex":3,"endColumnIndex":5},"cell":{"userEnteredFormat":{"backgroundColor":hc("#F4F9FD" if i%2==0 else "#FFFFFF")}},"fields":"userEnteredFormat(backgroundColor)"}})

    spreadsheet.batch_update({"requests":reqs})

    # チャート（円・棒）
    charts = [
        {"addChart":{"chart":{"spec":{"title":f"{m}月 費目別支出","titleTextFormat":{"bold":True},"pieChart":{"legendPosition":"RIGHT_LEGEND","threeDimensional":False,"series":{"sourceRange":{"sources":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+n,"startColumnIndex":1,"endColumnIndex":2}]}},"domain":{"sourceRange":{"sources":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+n,"startColumnIndex":0,"endColumnIndex":1}]}}}},"position":{"overlayPosition":{"anchorCell":{"sheetId":ws.id,"rowIndex":2,"columnIndex":5},"widthPixels":500,"heightPixels":340}}}}},
        {"addChart":{"chart":{"spec":{"title":f"{YEAR}年 月次支出推移","titleTextFormat":{"bold":True},"basicChart":{"chartType":"COLUMN","legendPosition":"NO_LEGEND","axis":[{"position":"BOTTOM_AXIS","title":""},{"position":"LEFT_AXIS","title":"支出（円）"}],"domains":[{"domain":{"sourceRange":{"sources":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+NUM_MONTHS,"startColumnIndex":3,"endColumnIndex":4}]}}}],"series":[{"series":{"sourceRange":{"sources":[{"sheetId":ws.id,"startRowIndex":10,"endRowIndex":10+NUM_MONTHS,"startColumnIndex":4,"endColumnIndex":5}]}},"targetAxis":"LEFT_AXIS","color":hc(BLUE)}]}},"position":{"overlayPosition":{"anchorCell":{"sheetId":ws.id,"rowIndex":30,"columnIndex":0},"widthPixels":700,"heightPixels":300}}}}},
    ]
    spreadsheet.batch_update({"requests":charts})
    print("  ✓ ダッシュボード")
    return ws


# ══════════════════════════════════════════════════════
# メイン
# ══════════════════════════════════════════════════════
def main():
    config = json.loads(CONFIG_PATH.read_text())
    spreadsheet_id = config["spreadsheet_id"]

    print("Google 認証中...")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(spreadsheet_id)

    # 新シート先作成
    existing = {ws.title for ws in spreadsheet.worksheets()}
    for title in ["ダッシュボード","収支管理","予算プラン","明細","月別集計"]:
        if title not in existing:
            spreadsheet.add_worksheet(title=title, rows=2000, cols=15)

    # 旧シート削除
    keep = {"ダッシュボード","収支管理","予算プラン","明細","月別集計"}
    for ws in spreadsheet.worksheets():
        if ws.title not in keep:
            spreadsheet.del_worksheet(ws)

    print("全シートを再構築中...")
    setup_meisai(spreadsheet)
    setup_monthly(spreadsheet)
    setup_budget_plan(spreadsheet)
    setup_shuushi(spreadsheet)
    setup_dashboard(spreadsheet)

    # シート順
    wmap = {ws.title: ws for ws in spreadsheet.worksheets()}
    for i, title in enumerate(["ダッシュボード","収支管理","予算プラン","明細","月別集計"]):
        if title in wmap:
            spreadsheet.batch_update({"requests":[{"updateSheetProperties":{"properties":{"sheetId":wmap[title].id,"index":i},"fields":"index"}}]})

    print(f"\n✅ 完了！")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")


if __name__ == "__main__":
    main()
