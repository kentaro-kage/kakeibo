#!/usr/bin/env python3
"""家計簿 Google スプレッドシート更新スクリプト"""

import json
import sys
from pathlib import Path
from datetime import datetime

import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
CREDENTIALS_PATH = SCRIPT_DIR / "credentials.json"
TOKEN_PATH = SCRIPT_DIR / "token.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

DAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]


def get_credentials():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                print(f"ERROR: {CREDENTIALS_PATH} が見つかりません。")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.write_text(creds.to_json())
    return creds


def append_receipt(data: dict):
    config = json.loads(CONFIG_PATH.read_text())
    spreadsheet_id = config.get("spreadsheet_id", "").strip()
    if not spreadsheet_id:
        print("ERROR: config.json の spreadsheet_id が空です。")
        sys.exit(1)

    creds = get_credentials()
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(spreadsheet_id)
    ws = spreadsheet.worksheet("明細")

    date_str = data["date"]
    d = datetime.strptime(date_str, "%Y-%m-%d")
    weekday = DAYS_JA[d.weekday()]
    store = data.get("store", "")
    memo = data.get("memo", "")

    rows = []
    for item in data.get("items", []):
        rows.append([
            date_str,
            weekday,
            store,
            item.get("name", ""),
            item.get("category", "その他"),
            item.get("amount", 0),
            memo,
        ])

    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")

    print(f"✓ 明細に {len(rows)} 件を追記しました。")
    for r in rows:
        print(f"  {r[0]}({r[1]}) | {r[2]} | {r[3]} | {r[4]} | ¥{int(r[5]):,}")


def main():
    raw = sys.argv[1] if len(sys.argv) >= 2 else sys.stdin.read()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"ERROR: JSON パース失敗: {e}")
        sys.exit(1)
    append_receipt(data)


if __name__ == "__main__":
    main()
