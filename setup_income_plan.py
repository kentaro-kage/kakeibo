#!/usr/bin/env python3
"""収支管理・予算プランシートのセットアップ"""

import json
from pathlib import Path
from datetime import date
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH  = SCRIPT_DIR / "config.json"
CREDENTIALS_PATH = SCRIPT_DIR / "credentials.json"
TOKEN_PATH   = SCRIPT_DIR / "token.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

YEAR = date.today().year
MONTH = date.today().month
MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]

NAVY   = "#1E3A5F"
GREEN  = "#1A5E39"
RED_D  = "#7B241C"
BLUE2  = "#1A5276"
WHITE_C = {"red":1,"green":1,"blue":1}

WIFE_INCOME    = 250000
HUSBAND_INCOME = 250000
TOTAL_INCOME   = WIFE_INCOME + HUSBAND_INCOME

CATEGORIES = [
    "家賃","光熱費","通信費","保育料","保険","外貨積立","税金",
    "医療費","衣料美容費","食費","日用品","健太郎お小遣い",
    "藍子お小遣い","陽和お小遣い","サブスク代","外食費","自動車","その他",
]

DEFAULT_BUDGETS = {
    "家賃":         150000,
    "光熱費":        15000,
    "通信費":        10000,
    "保育料":        45000,
    "保険":          25000,
    "外貨積立":      20000,
    "税金":          20000,
    "医療費":         5000,
    "衣料美容費":    15000,
    "食費":          50000,
    "日用品":        15000,
    "健太郎お小遣い": 20000,
    "藍子お小遣い":   20000,
    "陽和お小遣い":    3000,
    "サブスク代":     8000,
    "外食費":        20000,
    "自動車":        15000,
    "その他":        10000,
}


def hc(h):
    h = h.lstrip("#")
    return {"red":int(h[0:2],16)/255,"green":int(h[2:4],16)/255,"blue":int(h[4:6],16)/255}

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


# ════════════════════════════════════════════════
# 収支管理シート
# ════════════════════════════════════════════════
def setup_shuushi(spreadsheet):
    try:
        ws = spreadsheet.worksheet("収支管理")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="収支管理", rows=20, cols=16)

    # ── データ ──
    # タイトル行
    ws.update(values=[[f"{YEAR}年 月次収支管理"]], range_name="A1")

    # ヘッダー行
    ws.update(values=[["項目"] + MONTHS + ["年間合計"]], range_name="A2:N2")

    # 妻・夫収入（初期値入力済み・毎月編集可）
    wife_row    = ["妻（藍子）収入"] + [WIFE_INCOME]*12 + [f"=SUM(B3:M3)"]
    husband_row = ["夫（健太郎）収入"] + [HUSBAND_INCOME]*12 + [f"=SUM(B4:M4)"]
    ws.update(values=[wife_row], range_name="A3:N3")
    ws.update(values=[husband_row], range_name="A4:N4")

    # 収入合計
    income_total = ["収入合計"] + [f"=B3+B4" if i==0 else f"={chr(66+i)}3+{chr(66+i)}4" for i in range(12)] + ["=SUM(B5:M5)"]
    # Build column letters properly
    income_total_row = ["収入合計"]
    for i in range(12):
        col = chr(ord('B') + i)
        income_total_row.append(f"={col}3+{col}4")
    income_total_row.append("=SUM(B5:M5)")
    ws.update(values=[income_total_row], range_name="A5:N5", value_input_option="USER_ENTERED")

    # 空行
    ws.update(values=[["──────"]], range_name="A6")

    # 支出合計（明細から自動集計）
    expense_row = ["支出合計（実績）"]
    for mo in range(1, 13):
        expense_row.append(
            f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={mo})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)'
        )
    expense_row.append("=SUM(B7:M7)")
    ws.update(values=[expense_row], range_name="A7:N7", value_input_option="USER_ENTERED")

    # 収支差額
    diff_row = ["収支差額（月次）"]
    for i in range(12):
        col = chr(ord('B') + i)
        diff_row.append(f"={col}5-{col}7")
    diff_row.append("=SUM(B8:M8)")
    ws.update(values=[diff_row], range_name="A8:N8", value_input_option="USER_ENTERED")

    # 累計収支
    cum_row = ["累計収支"]
    cum_row.append("=B8")  # 1月
    for i in range(1, 12):
        prev_col = chr(ord('B') + i - 1)
        this_col = chr(ord('B') + i)
        cum_row.append(f"={prev_col}9+{this_col}8")
    cum_row.append("")
    ws.update(values=[cum_row], range_name="A9:N9", value_input_option="USER_ENTERED")

    # ── 書式 ──
    reqs = []

    # タイトル（マージ）
    reqs += [
        {"mergeCells": {
            "range": {"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":14},
            "mergeType":"MERGE_ALL"
        }},
        {"repeatCell": {
            "range": {"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":14},
            "cell": {"userEnteredFormat": {
                "backgroundColor": hc(NAVY),
                "textFormat": {"foregroundColor":WHITE_C,"bold":True,"fontSize":16},
                "horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"
            }},
            "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }},
        {"updateDimensionProperties": {
            "range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},
            "properties":{"pixelSize":56},"fields":"pixelSize"
        }},
    ]

    # ヘッダー行
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":1,"endRowIndex":2,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc(BLUE2),
            "textFormat": {"foregroundColor":WHITE_C,"bold":True,"fontSize":10},
            "horizontalAlignment":"CENTER"
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})

    # 妻行（ピンク系）
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":2,"endRowIndex":3,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc("#FDEBF7"),
            "textFormat": {"bold":True}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 夫行（青系）
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":3,"endRowIndex":4,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc("#D6EAF8"),
            "textFormat": {"bold":True}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 収入合計行（緑系）
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":4,"endRowIndex":5,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc("#D5F5E3"),
            "textFormat": {"bold":True,"fontSize":11}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 支出合計行（オレンジ系）
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":6,"endRowIndex":7,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc("#FDEBD0"),
            "textFormat": {"bold":True}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 収支差額行
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":7,"endRowIndex":8,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc(NAVY),
            "textFormat": {"foregroundColor":WHITE_C,"bold":True,"fontSize":11}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 累計収支行
    reqs.append({"repeatCell": {
        "range": {"sheetId":ws.id,"startRowIndex":8,"endRowIndex":9,"startColumnIndex":0,"endColumnIndex":14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hc("#1A5276"),
            "textFormat": {"foregroundColor":WHITE_C,"bold":True}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 数値書式（¥）：収入・支出・差額・累計
    for r_start, r_end in [(2,5),(6,10)]:
        reqs.append({"repeatCell": {
            "range": {"sheetId":ws.id,"startRowIndex":r_start,"endRowIndex":r_end,"startColumnIndex":1,"endColumnIndex":14},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type":"NUMBER","pattern":"¥#,##0"},
                "horizontalAlignment":"RIGHT"
            }},
            "fields":"userEnteredFormat(numberFormat,horizontalAlignment)"
        }})

    # 収支差額：マイナス赤（条件付き書式）
    reqs.append({"addConditionalFormatRule": {
        "rule": {
            "ranges": [{"sheetId":ws.id,"startRowIndex":7,"endRowIndex":8,"startColumnIndex":1,"endColumnIndex":13}],
            "booleanRule": {
                "condition": {"type":"NUMBER_LESS","values":[{"userEnteredValue":"0"}]},
                "format": {"textFormat":{"foregroundColor":hc("#E74C3C"),"bold":True}}
            }
        },
        "index":0
    }})

    # 列幅
    reqs.append({"updateDimensionProperties": {
        "range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":0,"endIndex":1},
        "properties":{"pixelSize":175},"fields":"pixelSize"
    }})
    for i in range(1,14):
        reqs.append({"updateDimensionProperties": {
            "range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},
            "properties":{"pixelSize":90},"fields":"pixelSize"
        }})

    # 行高
    reqs.append({"updateDimensionProperties": {
        "range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":1,"endIndex":10},
        "properties":{"pixelSize":36},"fields":"pixelSize"
    }})

    # 2行目まで固定
    reqs.append({"updateSheetProperties": {
        "properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":2}},
        "fields":"gridProperties.frozenRowCount"
    }})

    # ボーダー
    reqs.append({"updateBorders": {
        "range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":10,"startColumnIndex":0,"endColumnIndex":14},
        "innerHorizontal":{"style":"SOLID","color":hc("#D5D8DC")},
        "innerVertical":{"style":"SOLID","color":hc("#D5D8DC")},
    }})

    spreadsheet.batch_update({"requests": reqs})
    print("  ✓ 収支管理")
    return ws


# ════════════════════════════════════════════════
# 予算プランシート
# ════════════════════════════════════════════════
def setup_budget_plan(spreadsheet):
    try:
        ws = spreadsheet.worksheet("予算プラン")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="予算プラン", rows=50, cols=10)

    total_budget = sum(DEFAULT_BUDGETS.values())
    savings      = TOTAL_INCOME - total_budget

    # ── データ入力 ──

    # タイトル
    ws.update(values=[[f"{YEAR}年 家計予算プラン"]], range_name="A1")

    # 基本設定
    ws.update(values=[
        ["基本設定"],
        ["妻（藍子）月収",   WIFE_INCOME],
        ["夫（健太郎）月収", HUSBAND_INCOME],
        ["月収合計",         f"=B3+B4"],
        ["年収合計",         f"=B5*12"],
    ], range_name="A2:B6", value_input_option="USER_ENTERED")

    # 区切り
    ws.update(values=[["費目別 予算プラン"]], range_name="A8")

    # テーブルヘッダー
    ws.update(values=[["費目","月次予算","収入比率","今月実績","過不足","達成率","年間予算"]], range_name="A9:G9")

    # カテゴリ行
    n = len(CATEGORIES)
    rows = []
    for i, cat in enumerate(CATEGORIES):
        budget = DEFAULT_BUDGETS.get(cat, 0)
        row_num = 10 + i
        mo = MONTH
        actual_formula = (
            f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")'
            f'*(MONTH(明細!$A$2:$A$2000)={mo})'
            f'*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)'
        )
        rows.append([
            cat,
            budget,                                          # B: 月次予算（編集可）
            f"=B{row_num}/$B$5",                             # C: 収入比率
            actual_formula,                                  # D: 今月実績
            f"=B{row_num}-D{row_num}",                       # E: 過不足
            f'=IFERROR(D{row_num}/B{row_num},"")',            # F: 達成率
            f"=B{row_num}*12",                               # G: 年間予算
        ])
    ws.update(values=rows, range_name=f"A10:G{9+n}", value_input_option="USER_ENTERED")

    # 合計行
    sum_row = 10 + n
    ws.update(values=[[
        "支出合計",
        f"=SUM(B10:B{9+n})",
        f"=B{sum_row}/$B$5",
        f"=SUM(D10:D{9+n})",
        f"=B{sum_row}-D{sum_row}",
        f'=IFERROR(D{sum_row}/B{sum_row},"")',
        f"=B{sum_row}*12",
    ]], range_name=f"A{sum_row}:G{sum_row}", value_input_option="USER_ENTERED")

    # 貯蓄分析
    s = sum_row + 2
    ws.update(values=[["貯蓄分析"]], range_name=f"A{s}")
    ws.update(values=[
        ["月次収入合計",   "=$B$5"],
        ["月次支出予算",   f"=B{sum_row}"],
        ["月次貯蓄可能額", f"=$B$5-B{sum_row}"],
        ["貯蓄率",         f"=(B{s+3}/$B$5)"],
        ["年間貯蓄予測",   f"=B{s+3}*12"],
    ], range_name=f"A{s+1}:B{s+5}", value_input_option="USER_ENTERED")

    # ── 書式 ──
    reqs = []

    # タイトル（A1 マージ）
    reqs += [
        {"mergeCells": {
            "range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":7},
            "mergeType":"MERGE_ALL"
        }},
        {"repeatCell": {
            "range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":7},
            "cell":{"userEnteredFormat":{
                "backgroundColor":hc(NAVY),
                "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":16},
                "horizontalAlignment":"CENTER","verticalAlignment":"MIDDLE"
            }},
            "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }},
        {"updateDimensionProperties":{
            "range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},
            "properties":{"pixelSize":60},"fields":"pixelSize"
        }},
    ]

    # 基本設定セクション（A2:B6）
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":1,"endRowIndex":2,"startColumnIndex":0,"endColumnIndex":7},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#2E86C1"),
            "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":11},
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":2,"endRowIndex":6,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{"backgroundColor":hc("#EBF5FB")}},
        "fields":"userEnteredFormat(backgroundColor)"
    }})

    # 妻・夫収入 ハイライト
    for r, bg in [(2, "#FDEBF7"), (3, "#D6EAF8")]:
        reqs.append({"repeatCell": {
            "range":{"sheetId":ws.id,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":0,"endColumnIndex":2},
            "cell":{"userEnteredFormat":{"backgroundColor":hc(bg)}},
            "fields":"userEnteredFormat(backgroundColor)"
        }})

    # 費目別セクション見出し
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":7,"endRowIndex":8,"startColumnIndex":0,"endColumnIndex":7},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#1A5276"),
            "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":12},
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # テーブルヘッダー（A9:G9）
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":8,"endRowIndex":9,"startColumnIndex":0,"endColumnIndex":7},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc(NAVY),
            "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":10},
            "horizontalAlignment":"CENTER"
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})

    # カテゴリ行（交互）
    for i in range(n):
        bg = "#F4F9FD" if i%2==0 else "#FFFFFF"
        reqs.append({"repeatCell": {
            "range":{"sheetId":ws.id,"startRowIndex":9+i,"endRowIndex":10+i,"startColumnIndex":0,"endColumnIndex":7},
            "cell":{"userEnteredFormat":{"backgroundColor":hc(bg)}},
            "fields":"userEnteredFormat(backgroundColor)"
        }})

    # 合計行
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":9+n,"endRowIndex":10+n,"startColumnIndex":0,"endColumnIndex":7},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#2E86C1"),
            "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":11}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 過不足列：マイナス = 赤、プラス = 緑（条件付き書式）
    reqs.append({"addConditionalFormatRule": {
        "rule":{
            "ranges":[{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":9+n,"startColumnIndex":4,"endColumnIndex":5}],
            "booleanRule":{
                "condition":{"type":"NUMBER_LESS","values":[{"userEnteredValue":"0"}]},
                "format":{"textFormat":{"foregroundColor":hc("#E74C3C"),"bold":True},"backgroundColor":hc("#FDEDEC")}
            }
        },"index":0
    }})
    reqs.append({"addConditionalFormatRule": {
        "rule":{
            "ranges":[{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":9+n,"startColumnIndex":4,"endColumnIndex":5}],
            "booleanRule":{
                "condition":{"type":"NUMBER_GREATER_THAN_EQ","values":[{"userEnteredValue":"0"}]},
                "format":{"textFormat":{"foregroundColor":hc("#1A7431"),"bold":True},"backgroundColor":hc("#EAFAF1")}
            }
        },"index":1
    }})

    # 達成率 100%超 = 赤背景
    reqs.append({"addConditionalFormatRule": {
        "rule":{
            "ranges":[{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":9+n,"startColumnIndex":5,"endColumnIndex":6}],
            "booleanRule":{
                "condition":{"type":"NUMBER_GREATER","values":[{"userEnteredValue":"1"}]},
                "format":{"backgroundColor":hc("#FADBD8")}
            }
        },"index":2
    }})

    # 貯蓄分析セクション
    s_idx = sum_row  # 0-indexed: sum_row -1 ... let me recalculate
    s_row0 = 9 + n + 1  # 0-indexed row for "貯蓄分析" label (s = sum_row+2, 1-indexed)
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":s_row0,"endRowIndex":s_row0+1,"startColumnIndex":0,"endColumnIndex":7},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc(GREEN),
            "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":12}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})
    reqs.append({"repeatCell": {
        "range":{"sheetId":ws.id,"startRowIndex":s_row0+1,"endRowIndex":s_row0+6,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{"backgroundColor":hc("#EAFAF1")}},
        "fields":"userEnteredFormat(backgroundColor)"
    }})

    # 数値書式
    money_ranges = [
        (1,6,1,2), (8,9+n+1,1,3), (8,9+n+1,5,7),
        (s_row0+1, s_row0+5, 1, 2),
    ]
    for r0,r1,c0,c1 in money_ranges:
        reqs.append({"repeatCell":{
            "range":{"sheetId":ws.id,"startRowIndex":r0,"endRowIndex":r1,"startColumnIndex":c0,"endColumnIndex":c1},
            "cell":{"userEnteredFormat":{
                "numberFormat":{"type":"NUMBER","pattern":"¥#,##0"},
                "horizontalAlignment":"RIGHT"
            }},
            "fields":"userEnteredFormat(numberFormat,horizontalAlignment)"
        }})

    # 収入比率・達成率 → %書式
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":9+n+1,"startColumnIndex":2,"endColumnIndex":3},
        "cell":{"userEnteredFormat":{"numberFormat":{"type":"PERCENT","pattern":"0.0%"},"horizontalAlignment":"CENTER"}},
        "fields":"userEnteredFormat(numberFormat,horizontalAlignment)"
    }})
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":9,"endRowIndex":9+n+1,"startColumnIndex":5,"endColumnIndex":6},
        "cell":{"userEnteredFormat":{"numberFormat":{"type":"PERCENT","pattern":"0%"},"horizontalAlignment":"CENTER"}},
        "fields":"userEnteredFormat(numberFormat,horizontalAlignment)"
    }})
    # 基本設定 年収
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":4,"endRowIndex":6,"startColumnIndex":1,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"}}},
        "fields":"userEnteredFormat(numberFormat)"
    }})
    # 貯蓄率 %
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":s_row0+4,"endRowIndex":s_row0+5,"startColumnIndex":1,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{"numberFormat":{"type":"PERCENT","pattern":"0.0%"}}},
        "fields":"userEnteredFormat(numberFormat)"
    }})

    # 列幅
    for ci, px in enumerate([165, 110, 80, 110, 90, 70, 100]):
        reqs.append({"updateDimensionProperties":{
            "range":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":ci,"endIndex":ci+1},
            "properties":{"pixelSize":px},"fields":"pixelSize"
        }})

    # 行高
    reqs.append({"updateDimensionProperties":{
        "range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":9+n+8},
        "properties":{"pixelSize":32},"fields":"pixelSize"
    }})
    reqs.append({"updateDimensionProperties":{
        "range":{"sheetId":ws.id,"dimension":"ROWS","startIndex":0,"endIndex":1},
        "properties":{"pixelSize":60},"fields":"pixelSize"
    }})

    # 1行目固定
    reqs.append({"updateSheetProperties":{
        "properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":9}},
        "fields":"gridProperties.frozenRowCount"
    }})

    # ボーダー
    reqs.append({"updateBorders":{
        "range":{"sheetId":ws.id,"startRowIndex":8,"endRowIndex":9+n+1,"startColumnIndex":0,"endColumnIndex":7},
        "innerHorizontal":{"style":"SOLID","color":hc("#D5D8DC")},
        "innerVertical":{"style":"SOLID","color":hc("#D5D8DC")},
        "top":{"style":"SOLID_MEDIUM","color":hc("#2E86C1")},
        "bottom":{"style":"SOLID_MEDIUM","color":hc("#2E86C1")},
        "left":{"style":"SOLID_MEDIUM","color":hc("#2E86C1")},
        "right":{"style":"SOLID_MEDIUM","color":hc("#2E86C1")},
    }})

    spreadsheet.batch_update({"requests": reqs})
    print("  ✓ 予算プラン")
    return ws


# ════════════════════════════════════════════════
# ダッシュボード収支エリア更新
# ════════════════════════════════════════════════
def update_dashboard(spreadsheet):
    ws = spreadsheet.worksheet("ダッシュボード")

    # 既存の収支サマリーを書き直し（A3:B7 エリアを拡張）
    ws.update(values=[
        ["📅 今月",         f"{MONTH}月"],
        ["💸 今月の支出",   f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={MONTH})*(YEAR(明細!$A$2:$A$2000)={YEAR})*明細!$F$2:$F$2000)'],
        ["💰 今月の収入",   f'=収支管理!{chr(ord("B")+MONTH-1)}5'],
        ["📊 収支差額",     f'=B5-B4'],
        ["📦 今月の件数",   f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={MONTH})*(YEAR(明細!$A$2:$A$2000)={YEAR})*(明細!$A$2:$A$2000<>""))'],
    ], range_name="A3:B7", value_input_option="USER_ENTERED")

    reqs = []

    # 今月ヘッダー
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":2,"endRowIndex":3,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#2E86C1"),
            "textFormat":{"foregroundColor":WHITE_C,"bold":True,"fontSize":11},
            "horizontalAlignment":"CENTER"
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})

    # 支出（赤）
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":3,"endRowIndex":4,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#FDEDEC"),
            "textFormat":{"bold":True,"fontSize":13,"foregroundColor":hc("#C0392B")}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 収入（緑）
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":4,"endRowIndex":5,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#EAFAF1"),
            "textFormat":{"bold":True,"fontSize":13,"foregroundColor":hc("#1A7431")}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 収支差額（紺）
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":5,"endRowIndex":6,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc(NAVY),
            "textFormat":{"bold":True,"fontSize":13,"foregroundColor":WHITE_C}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 件数
    reqs.append({"repeatCell":{
        "range":{"sheetId":ws.id,"startRowIndex":6,"endRowIndex":7,"startColumnIndex":0,"endColumnIndex":2},
        "cell":{"userEnteredFormat":{
            "backgroundColor":hc("#EBF5FB"),
            "textFormat":{"bold":False}
        }},
        "fields":"userEnteredFormat(backgroundColor,textFormat)"
    }})

    # 数値書式
    for r in [3,4,5]:
        reqs.append({"repeatCell":{
            "range":{"sheetId":ws.id,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":1,"endColumnIndex":2},
            "cell":{"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"¥#,##0"}}},
            "fields":"userEnteredFormat(numberFormat)"
        }})

    # 収支差額: マイナス赤（条件付き書式）
    reqs.append({"addConditionalFormatRule":{
        "rule":{
            "ranges":[{"sheetId":ws.id,"startRowIndex":5,"endRowIndex":6,"startColumnIndex":1,"endColumnIndex":2}],
            "booleanRule":{
                "condition":{"type":"NUMBER_LESS","values":[{"userEnteredValue":"0"}]},
                "format":{"backgroundColor":hc("#FADBD8"),"textFormat":{"foregroundColor":hc("#C0392B"),"bold":True}}
            }
        },"index":0
    }})

    spreadsheet.batch_update({"requests": reqs})
    print("  ✓ ダッシュボード更新")


# ════════════════════════════════════════════════
# メイン
# ════════════════════════════════════════════════
def main():
    config = json.loads(CONFIG_PATH.read_text())
    spreadsheet_id = config["spreadsheet_id"]

    print("Google 認証中...")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(spreadsheet_id)

    # 新シートを先に作成
    existing = {ws.title for ws in spreadsheet.worksheets()}
    for title in ["収支管理","予算プラン"]:
        if title not in existing:
            spreadsheet.add_worksheet(title=title, rows=60, cols=14)

    print("シートをセットアップ中...")
    setup_shuushi(spreadsheet)
    setup_budget_plan(spreadsheet)
    update_dashboard(spreadsheet)

    # シート順：ダッシュボード→収支管理→予算プラン→明細→月別集計
    wmap = {ws.title: ws for ws in spreadsheet.worksheets()}
    order = ["ダッシュボード","収支管理","予算プラン","明細","月別集計"]
    for i, title in enumerate(order):
        if title in wmap:
            spreadsheet.batch_update({"requests":[{"updateSheetProperties":{
                "properties":{"sheetId":wmap[title].id,"index":i},
                "fields":"index"
            }}]})

    print(f"\n✅ 完了！")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")


if __name__ == "__main__":
    main()
