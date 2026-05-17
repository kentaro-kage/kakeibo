#!/usr/bin/env python3
"""LINE Bot webhook server — 家計簿自動記録"""

import os
import base64
import json
from datetime import datetime

import anthropic
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import PlainTextResponse
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi, MessagingApiBlob,
    ReplyMessageRequest, TextMessage,
)
from linebot.v3.webhooks import (
    MessageEvent, ImageMessageContent, TextMessageContent,
)
from linebot.v3.exceptions import InvalidSignatureError
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request as GoogleRequest

# ══════════════════════════════════════════════
# 設定（環境変数から取得）
# ══════════════════════════════════════════════
LINE_CHANNEL_SECRET       = os.environ["LINE_CHANNEL_SECRET"]
LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
ANTHROPIC_API_KEY         = os.environ["ANTHROPIC_API_KEY"]
SPREADSHEET_ID            = os.environ["SPREADSHEET_ID"]
GOOGLE_TOKEN_B64          = os.environ["GOOGLE_TOKEN_B64"]

SCOPES  = ["https://www.googleapis.com/auth/spreadsheets"]
DAYS_JA = ["月","火","水","木","金","土","日"]

CATEGORIES = [
    "家賃","光熱費","通信費","保育料","保険","外貨積立","税金",
    "医療費","衣料美容費","食費","日用品","健太郎お小遣い",
    "藍子お小遣い","陽和お小遣い","サブスク代","外食費","自動車","ペット","その他",
]

RULES = """
- お酒・アルコール類 → 健太郎お小遣い
- ガソリン → 自動車
- 損保ジャパン → 自動車
- コストコのレシートは品目ごとに分類（食品→食費、日用品→日用品、ガソリン→自動車）
- 薬・医薬品 → 医療費
- コーヒー・食材など家庭用食品 → 食費
- 外食・テイクアウト → 外食費
"""

app = FastAPI()
handler    = WebhookHandler(LINE_CHANNEL_SECRET)
line_conf  = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)


# ══════════════════════════════════════════════
# Google Sheets
# ══════════════════════════════════════════════
def get_worksheet():
    token_info = json.loads(base64.b64decode(GOOGLE_TOKEN_B64).decode())
    creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(GoogleRequest())
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet("明細")


def append_to_sheet(data: dict) -> list:
    ws = get_worksheet()
    date_str = data["date"]
    d        = datetime.strptime(date_str, "%Y-%m-%d")
    weekday  = DAYS_JA[d.weekday()]
    store    = data.get("store", "")
    memo     = data.get("memo", "")

    rows = []
    for item in data.get("items", []):
        rows.append([
            date_str, weekday, store,
            item.get("name", ""),
            item.get("category", "その他"),
            item.get("amount", 0),
            memo,
        ])
    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
    return rows


# ══════════════════════════════════════════════
# Claude API — レシート画像解析
# ══════════════════════════════════════════════
def analyze_image(image_bytes: bytes) -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    today  = datetime.now().strftime("%Y-%m-%d")
    cats   = "、".join(CATEGORIES)
    img_b64 = base64.standard_b64encode(image_bytes).decode()

    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64},
                },
                {
                    "type": "text",
                    "text": f"""このレシート画像を分析してJSON形式のみ返してください。

{{
  "date": "YYYY-MM-DD",
  "store": "店舗名",
  "items": [{{"name": "品目名", "category": "費目", "amount": 税込金額(整数)}}],
  "memo": ""
}}

費目（必ずこの中から選ぶ）: {cats}

分類ルール:
{RULES}

日付が読み取れない場合は今日({today})を使用。
JSONのみ返答（説明文・コードブロック不要）。""",
                },
            ],
        }],
    )

    text = resp.content[0].text.strip()
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text.strip())


# ══════════════════════════════════════════════
# Claude API — テキストメッセージ解析
# ══════════════════════════════════════════════
def analyze_text(text: str) -> dict | None:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    today  = datetime.now().strftime("%Y-%m-%d")
    cats   = "、".join(CATEGORIES)

    resp = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=512,
        messages=[{
            "role": "user",
            "content": f"""以下のメッセージが家計の支出情報であればJSONを、違えばnullを返してください。

メッセージ:「{text}」

支出の場合の出力:
{{"date":"YYYY-MM-DD","store":"","items":[{{"name":"品目","category":"費目","amount":金額}}],"memo":""}}

費目: {cats}
分類ルール: {RULES}
今日: {today}

JSONまたはnullのみ返答。""",
        }],
    )

    t = resp.content[0].text.strip()
    if t == "null" or not t:
        return None
    if "```" in t:
        t = t.split("```")[1]
        if t.startswith("json"):
            t = t[4:]
    try:
        return json.loads(t.strip())
    except Exception:
        return None


# ══════════════════════════════════════════════
# LINE Webhook
# ══════════════════════════════════════════════
@app.post("/callback")
async def callback(request: Request):
    sig  = request.headers.get("X-Line-Signature", "")
    body = await request.body()
    try:
        handler.handle(body.decode(), sig)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")
    return PlainTextResponse("OK")


@app.get("/health")
def health():
    return {"status": "ok"}


@handler.add(MessageEvent, message=ImageMessageContent)
def on_image(event):
    with ApiClient(line_conf) as api_client:
        bot  = MessagingApi(api_client)
        blob = MessagingApiBlob(api_client)
        try:
            image_bytes = blob.get_message_content(event.message.id)
            data  = analyze_image(image_bytes)
            rows  = append_to_sheet(data)

            items_str = "\n".join(
                [f"  {r[3]}（{r[4]}）¥{int(r[5]):,}" for r in rows]
            )
            total = sum(int(r[5]) for r in rows)
            reply = (
                f"✅ {data['date']} {data.get('store','')} を記録しました\n"
                f"{items_str}\n"
                f"合計: ¥{total:,}"
            )
        except Exception as e:
            reply = f"⚠️ 読み取りエラー: {str(e)[:120]}"

        bot.reply_message(ReplyMessageRequest(
            reply_token=event.reply_token,
            messages=[TextMessage(text=reply)],
        ))


@handler.add(MessageEvent, message=TextMessageContent)
def on_text(event):
    text = event.message.text
    if len(text) < 3:
        return
    with ApiClient(line_conf) as api_client:
        bot = MessagingApi(api_client)
        try:
            data = analyze_text(text)
            if not data:
                return
            rows  = append_to_sheet(data)
            items_str = "\n".join(
                [f"  {r[3]}（{r[4]}）¥{int(r[5]):,}" for r in rows]
            )
            bot.reply_message(ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=f"✅ 記録しました\n{items_str}")],
            ))
        except Exception:
            pass  # 家計と無関係なテキストは無視


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
