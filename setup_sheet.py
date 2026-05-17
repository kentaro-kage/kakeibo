#!/usr/bin/env python3
"""家計簿スプレッドシート 初期セットアップ"""

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

CATEGORIES = [
    "家賃", "光熱費", "通信費", "保育料", "保険", "外貨積立", "税金",
    "医療費", "衣料美容費", "食費", "日用品", "健太郎お小遣い",
    "藍子お小遣い", "陽和お小遣い", "サブスク代", "外食費", "自動車", "その他"
]

# Pastel colors for row highlighting
CATEGORY_COLORS = {
    "家賃":        "#FADBD8",
    "光熱費":      "#FDEBD0",
    "通信費":      "#D6EAF8",
    "保育料":      "#FDEDEC",
    "保険":        "#E8DAEF",
    "外貨積立":    "#D1F2EB",
    "税金":        "#EAECEE",
    "医療費":      "#F9EBEA",
    "衣料美容費":  "#FDEBF7",
    "食費":        "#D5F5E3",
    "日用品":      "#EAFAF1",
    "健太郎お小遣い": "#DAE8FC",
    "藍子お小遣い":   "#E8D5FC",
    "陽和お小遣い":   "#FEF9E7",
    "サブスク代":  "#D5D8DC",
    "外食費":      "#FDEBD0",
    "自動車":      "#D6DBDF",
    "その他":      "#F2F3F4",
}

MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
NAVY   = "#1E3A5F"
WHITE  = {"red": 1, "green": 1, "blue": 1}


def hex_to_color(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2],16)/255, "green": int(h[2:4],16)/255, "blue": int(h[4:6],16)/255}


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


# ──────────────────────────────────────────────
# 明細シート
# ──────────────────────────────────────────────
def setup_meisai(spreadsheet):
    try:
        ws = spreadsheet.worksheet("明細")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="明細", rows=2000, cols=7)

    ws.update("A1:G1", [["日付", "曜日", "店舗名", "品目", "費目", "金額(税込)", "メモ"]])

    reqs = []

    # ヘッダー書式
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color(NAVY),
            "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 11},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
    }})

    # ヘッダー行高
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": ws.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 40}, "fields": "pixelSize"
    }})

    # 1行目固定
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount"
    }})

    # 列幅
    for i, px in enumerate([100, 50, 150, 150, 130, 110, 220]):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"
        }})

    # 金額列 数値書式
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2000, "startColumnIndex": 5, "endColumnIndex": 6},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "¥#,##0"},
            "horizontalAlignment": "RIGHT",
        }},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
    }})

    # 費目別 行カラー（条件付き書式）
    for idx, cat in enumerate(CATEGORIES):
        reqs.append({"addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": ws.id, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 7}],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f'=$E2="{cat}"'}]},
                    "format": {"backgroundColor": hex_to_color(CATEGORY_COLORS[cat])}
                }
            },
            "index": idx
        }})

    # 外枠ボーダー
    reqs.append({"updateBorders": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 7},
        "bottom": {"style": "SOLID_MEDIUM", "color": hex_to_color("#FFFFFF")}
    }})

    spreadsheet.batch_update({"requests": reqs})
    print("  ✓ 明細")
    return ws


# ──────────────────────────────────────────────
# 月別集計シート
# ──────────────────────────────────────────────
def setup_monthly(spreadsheet):
    year = date.today().year
    try:
        ws = spreadsheet.worksheet("月別集計")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="月別集計", rows=30, cols=15)

    # ヘッダー行
    ws.update("A1:N1", [["費目"] + MONTHS + ["年間合計"]])

    # 費目行（SUMPRODUCT式）
    rows = []
    for cat in CATEGORIES:
        row = [cat]
        for m in range(1, 13):
            row.append(
                f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")'
                f'*(MONTH(明細!$A$2:$A$2000)={m})'
                f'*(YEAR(明細!$A$2:$A$2000)={year})'
                f'*明細!$F$2:$F$2000)'
            )
        row.append(f'=SUM(B{len(rows)+2}:M{len(rows)+2})')
        rows.append(row)

    # 合計行
    total = ["合計"]
    for i, col in enumerate("BCDEFGHIJKLM"):
        total.append(f'=SUM({col}2:{col}{len(CATEGORIES)+1})')
    total.append(f'=SUM(N2:N{len(CATEGORIES)+1})')
    rows.append(total)

    ws.update(f"A2:N{len(CATEGORIES)+2}", rows, value_input_option="USER_ENTERED")

    n = len(CATEGORIES)
    reqs = []

    # ヘッダー
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color(NAVY),
            "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
    }})

    # 費目列
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": n+1, "startColumnIndex": 0, "endColumnIndex": 1},
        "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True},
            "backgroundColor": hex_to_color("#EBF5FB")
        }},
        "fields": "userEnteredFormat(textFormat,backgroundColor)"
    }})

    # 合計行
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": n+1, "endRowIndex": n+2, "startColumnIndex": 0, "endColumnIndex": 14},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color("#2E86C1"),
            "textFormat": {"foregroundColor": WHITE, "bold": True},
            "horizontalAlignment": "CENTER"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})

    # 数値書式
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": n+2, "startColumnIndex": 1, "endColumnIndex": 14},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "¥#,##0"},
            "horizontalAlignment": "RIGHT"
        }},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
    }})

    # 行高
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": ws.id, "dimension": "ROWS", "startIndex": 0, "endIndex": n+2},
        "properties": {"pixelSize": 34}, "fields": "pixelSize"
    }})

    # 固定
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 1}},
        "fields": "gridProperties(frozenRowCount,frozenColumnCount)"
    }})

    # 列幅
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 165}, "fields": "pixelSize"
    }})
    for i in range(1, 14):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1},
            "properties": {"pixelSize": 88}, "fields": "pixelSize"
        }})

    # 交互行カラー
    for i in range(n):
        bg = "#F4F9FD" if i % 2 == 0 else "#FFFFFF"
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": i+1, "endRowIndex": i+2, "startColumnIndex": 1, "endColumnIndex": 14},
            "cell": {"userEnteredFormat": {"backgroundColor": hex_to_color(bg)}},
            "fields": "userEnteredFormat(backgroundColor)"
        }})

    # ボーダー
    reqs.append({"updateBorders": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": n+2, "startColumnIndex": 0, "endColumnIndex": 14},
        "innerHorizontal": {"style": "SOLID", "color": hex_to_color("#D5D8DC")},
        "innerVertical": {"style": "SOLID", "color": hex_to_color("#D5D8DC")},
        "top":    {"style": "SOLID_MEDIUM", "color": hex_to_color("#2E86C1")},
        "bottom": {"style": "SOLID_MEDIUM", "color": hex_to_color("#2E86C1")},
        "left":   {"style": "SOLID_MEDIUM", "color": hex_to_color("#2E86C1")},
        "right":  {"style": "SOLID_MEDIUM", "color": hex_to_color("#2E86C1")},
    }})

    spreadsheet.batch_update({"requests": reqs})
    print("  ✓ 月別集計")
    return ws


# ──────────────────────────────────────────────
# ダッシュボードシート
# ──────────────────────────────────────────────
def setup_dashboard(spreadsheet):
    year = date.today().year
    m = date.today().month

    try:
        ws = spreadsheet.worksheet("ダッシュボード")
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="ダッシュボード", rows=60, cols=12)

    # ── データ配置 ──
    # A1: タイトル（後でマージ）
    ws.update("A1", [[f"{year}年 家計簿ダッシュボード"]])

    # 今月サマリー
    ws.update("A3:B3", [["📅 今月", f"{m}月"]])
    ws.update("A4:B4", [["💸 今月の支出合計",
        f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={year})*明細!$F$2:$F$2000)']])
    ws.update("A5:B5", [["📊 今月の件数",
        f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={year})*(明細!$A$2:$A$2000<>""))']])
    ws.update("A6:B6", [["📈 今月の平均/件",
        f'=IFERROR(B4/B5,0)']])

    # 費目別集計（チャート用データ）
    ws.update("A8:B8", [["費目", f"{m}月支出"]])
    cat_rows = []
    for i, cat in enumerate(CATEGORIES):
        cat_rows.append([cat,
            f'=SUMPRODUCT((明細!$E$2:$E$2000="{cat}")*(MONTH(明細!$A$2:$A$2000)={m})*(YEAR(明細!$A$2:$A$2000)={year})*明細!$F$2:$F$2000)'
        ])
    ws.update(f"A9:B{8+len(CATEGORIES)}", cat_rows, value_input_option="USER_ENTERED")

    # 月次推移（チャート用）
    ws.update("D8:E8", [["月", "支出合計"]])
    month_rows = []
    for mo in range(1, 13):
        month_rows.append([f"{mo}月",
            f'=SUMPRODUCT((MONTH(明細!$A$2:$A$2000)={mo})*(YEAR(明細!$A$2:$A$2000)={year})*明細!$F$2:$F$2000)'
        ])
    ws.update(f"D9:E20", month_rows, value_input_option="USER_ENTERED")

    n = len(CATEGORIES)
    reqs = []

    # タイトル行（マージ）
    reqs.append({"mergeCells": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 12},
        "mergeType": "MERGE_ALL"
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 12},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color(NAVY),
            "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 18},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
    }})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": ws.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 64}, "fields": "pixelSize"
    }})

    # サマリー カード（A3:B6）
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color("#2E86C1"),
            "textFormat": {"foregroundColor": WHITE, "bold": True, "fontSize": 11},
            "horizontalAlignment": "CENTER"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})
    for row_i, (bg, fs, bold) in enumerate(
        [("#EBF5FB", 11, True), ("#D6EAF8", 14, True), ("#EBF5FB", 11, False)], start=3
    ):
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": row_i, "endRowIndex": row_i+1, "startColumnIndex": 0, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {
                "backgroundColor": hex_to_color(bg),
                "textFormat": {"bold": bold, "fontSize": fs},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }})
    # 支出合計 赤字・大フォント
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": 4, "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "¥#,##0"},
            "textFormat": {"bold": True, "fontSize": 16, "foregroundColor": hex_to_color("#C0392B")}
        }},
        "fields": "userEnteredFormat(numberFormat,textFormat)"
    }})
    # 件数・平均
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 4, "endRowIndex": 6, "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"},
            "horizontalAlignment": "RIGHT"
        }},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 5, "endRowIndex": 6, "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "¥#,##0"}
        }},
        "fields": "userEnteredFormat(numberFormat)"
    }})

    # 費目テーブル ヘッダー（A8:B8）
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 7, "endRowIndex": 8, "startColumnIndex": 0, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color("#1A5276"),
            "textFormat": {"foregroundColor": WHITE, "bold": True},
            "horizontalAlignment": "CENTER"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})
    # 費目テーブル データ
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 8, "endRowIndex": 8+n, "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "¥#,##0"},
            "horizontalAlignment": "RIGHT"
        }},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
    }})
    for i in range(n):
        bg = "#F4F9FD" if i % 2 == 0 else "#FFFFFF"
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 8+i, "endRowIndex": 9+i, "startColumnIndex": 0, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {"backgroundColor": hex_to_color(bg)}},
            "fields": "userEnteredFormat(backgroundColor)"
        }})

    # 月次推移テーブル ヘッダー（D8:E8）
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 7, "endRowIndex": 8, "startColumnIndex": 3, "endColumnIndex": 5},
        "cell": {"userEnteredFormat": {
            "backgroundColor": hex_to_color("#1A5276"),
            "textFormat": {"foregroundColor": WHITE, "bold": True},
            "horizontalAlignment": "CENTER"
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 8, "endRowIndex": 20, "startColumnIndex": 4, "endColumnIndex": 5},
        "cell": {"userEnteredFormat": {
            "numberFormat": {"type": "NUMBER", "pattern": "¥#,##0"},
            "horizontalAlignment": "RIGHT"
        }},
        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
    }})
    for i in range(12):
        bg = "#F4F9FD" if i % 2 == 0 else "#FFFFFF"
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 8+i, "endRowIndex": 9+i, "startColumnIndex": 3, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": {"backgroundColor": hex_to_color(bg)}},
            "fields": "userEnteredFormat(backgroundColor)"
        }})

    # 列幅
    for ci, px in enumerate([175, 120, 20, 70, 110]):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": ci, "endIndex": ci+1},
            "properties": {"pixelSize": px}, "fields": "pixelSize"
        }})

    spreadsheet.batch_update({"requests": reqs})

    # ── チャート ──
    charts = []

    # 円グラフ（費目別）
    charts.append({"addChart": {"chart": {
        "spec": {
            "title": f"{m}月 費目別支出",
            "titleTextFormat": {"bold": True, "fontSize": 12},
            "pieChart": {
                "legendPosition": "RIGHT_LEGEND",
                "threeDimensional": False,
                "series": {"sourceRange": {"sources": [{
                    "sheetId": ws.id, "startRowIndex": 8, "endRowIndex": 8+n,
                    "startColumnIndex": 1, "endColumnIndex": 2
                }]}},
                "domain": {"sourceRange": {"sources": [{
                    "sheetId": ws.id, "startRowIndex": 8, "endRowIndex": 8+n,
                    "startColumnIndex": 0, "endColumnIndex": 1
                }]}}
            }
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": ws.id, "rowIndex": 2, "columnIndex": 5},
            "widthPixels": 500, "heightPixels": 340
        }}
    }}})

    # 棒グラフ（月次推移）
    charts.append({"addChart": {"chart": {
        "spec": {
            "title": f"{year}年 月次支出推移",
            "titleTextFormat": {"bold": True, "fontSize": 12},
            "basicChart": {
                "chartType": "COLUMN",
                "legendPosition": "NO_LEGEND",
                "axis": [
                    {"position": "BOTTOM_AXIS", "title": ""},
                    {"position": "LEFT_AXIS", "title": "支出（円）"}
                ],
                "domains": [{"domain": {"sourceRange": {"sources": [{
                    "sheetId": ws.id, "startRowIndex": 8, "endRowIndex": 20,
                    "startColumnIndex": 3, "endColumnIndex": 4
                }]}}}],
                "series": [{"series": {"sourceRange": {"sources": [{
                    "sheetId": ws.id, "startRowIndex": 8, "endRowIndex": 20,
                    "startColumnIndex": 4, "endColumnIndex": 5
                }]}}, "targetAxis": "LEFT_AXIS",
                "color": hex_to_color("#2E86C1")}]
            }
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": ws.id, "rowIndex": 28, "columnIndex": 0},
            "widthPixels": 700, "heightPixels": 320
        }}
    }}})

    spreadsheet.batch_update({"requests": charts})
    print("  ✓ ダッシュボード")
    return ws


# ──────────────────────────────────────────────
# メイン
# ──────────────────────────────────────────────
def main():
    config = json.loads(CONFIG_PATH.read_text())
    spreadsheet_id = config["spreadsheet_id"]

    print("Google 認証中...")
    creds = get_credentials()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(spreadsheet_id)

    print("シートをセットアップ中...")

    # 先に新シートを作成（存在しない場合のみ）
    existing = {ws.title for ws in spreadsheet.worksheets()}
    for title in ["明細", "月別集計", "ダッシュボード"]:
        if title not in existing:
            spreadsheet.add_worksheet(title=title, rows=2000, cols=15)

    # 不要シートを削除（新シートを作った後に行う）
    keep = {"明細", "月別集計", "ダッシュボード"}
    for ws in spreadsheet.worksheets():
        if ws.title not in keep:
            spreadsheet.del_worksheet(ws)

    setup_meisai(spreadsheet)
    setup_monthly(spreadsheet)
    setup_dashboard(spreadsheet)

    # シート順を並び替え: ダッシュボード → 明細 → 月別集計
    wmap = {ws.title: ws for ws in spreadsheet.worksheets()}
    for i, title in enumerate(["ダッシュボード", "明細", "月別集計"]):
        if title in wmap:
            spreadsheet.batch_update({"requests": [{"updateSheetProperties": {
                "properties": {"sheetId": wmap[title].id, "index": i},
                "fields": "index"
            }}]})

    print(f"\n✅ 完了！")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")


if __name__ == "__main__":
    main()
