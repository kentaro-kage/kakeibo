#!/usr/bin/env python3
"""家計簿入力アプリ（Streamlit）"""

import json
from pathlib import Path
from datetime import date, datetime

import gspread
import streamlit as st
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# ── 設定 ──────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "config.json"
TOKEN_PATH  = SCRIPT_DIR / "token.json"
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets"]
DAYS_JA     = ["月","火","水","木","金","土","日"]

config     = json.loads(CONFIG_PATH.read_text())
CATEGORIES = config["categories"]
SPREAD_ID  = config["spreadsheet_id"]

# ── Google Sheets ──────────────────────────────
@st.cache_resource
def get_gc():
    # Streamlit Cloud: secretsからtoken.jsonを生成
    if "google_token" in st.secrets:
        import tempfile
        token_data = dict(st.secrets["google_token"])
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
        json.dump(token_data, tmp)
        tmp.flush()
        creds = Credentials.from_authorized_user_file(tmp.name, SCOPES)
    else:
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return gspread.authorize(creds)

WATCH_BUDGETS = {
    "食費":           30000,
    "外食費":         15000,
    "健太郎お小遣い": 10000,
    "藍子お小遣い":   10000,
    "陽和お小遣い":   10000,
}

WATCH_GROUPS = [
    {
        "label": "食費",
        "items": ["食費", "外食費"],
    },
    {
        "label": "お小遣い",
        "items": ["健太郎お小遣い", "藍子お小遣い", "陽和お小遣い"],
    },
]

@st.cache_data(ttl=60)
def get_monthly_spending(year: int, month: int):
    """今月の費目別合計を返す。エラー時はメッセージも返す。"""
    try:
        gc = get_gc()
        ws = gc.open_by_key(SPREAD_ID).worksheet("明細")
        rows = ws.get_all_values()[1:]  # ヘッダー除く
        # 日付フォーマットを両方対応（2026-05 / 2026/05）
        prefix_h = f"{year}-{month:02d}"
        prefix_s = f"{year}/{month:02d}"
        totals = {k: 0 for k in WATCH_BUDGETS}
        for r in rows:
            if len(r) >= 6 and (r[0].startswith(prefix_h) or r[0].startswith(prefix_s)):
                cat = r[4]  # 費目列
                if cat in totals:
                    try:
                        totals[cat] += int(float(str(r[5]).replace(",", "")))
                    except ValueError:
                        pass
        return totals, None
    except Exception as e:
        return {k: 0 for k in WATCH_BUDGETS}, str(e)

def append_to_sheet(entries: list):
    gc = get_gc()
    ws = gc.open_by_key(SPREAD_ID).worksheet("明細")
    rows = []
    for e in entries:
        d = datetime.strptime(e["date"], "%Y-%m-%d")
        rows.append([e["date"], DAYS_JA[d.weekday()], e["note"], e["category"], e["category"], e["amount"], ""])
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    return rows

# ── ページ設定 ─────────────────────────────────
st.set_page_config(
    page_title="家計簿",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+JP:wght@300;400;500;700&display=swap');

/* ─── リセット & ベース ─────────────────────── */
#MainMenu, footer, header { visibility: hidden; height: 0; }
[data-testid="stToolbar"] { display: none; }
[data-testid="stDecoration"] { display: none; }
[data-testid="stHeader"] { display: none; }

html, body, [class*="css"] {
    font-family: 'Inter', 'Noto Sans JP', -apple-system, sans-serif;
}

.main .block-container {
    padding: 1.2rem 1rem 4rem;
    max-width: 430px;
}

/* ─── ヘッダー ──────────────────────────────── */
.page-header {
    padding: 0.8rem 0 1.2rem;
    border-bottom: 1px solid rgba(150,150,150,0.2);
    margin-bottom: 1.2rem;
}
.page-header-label {
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #9CA3AF;
    margin-bottom: 0.4rem;
}
.page-header-title {
    font-size: 1.75rem;
    font-weight: 700;
    color: inherit;
    letter-spacing: -0.03em;
    line-height: 1.1;
}
.page-header-date {
    font-size: 0.82rem;
    color: #9CA3AF;
    margin-top: 0.35rem;
    font-weight: 400;
}

/* ─── カード ────────────────────────────────── */
.entry-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 1.3rem 1.3rem 0.6rem;
    margin-bottom: 0.75rem;
    border: 1px solid #E5E7EB;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03);
}
.card-index {
    font-size: 0.68rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #9CA3AF;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 0.4rem;
}
.card-index::after {
    content: '';
    flex: 1;
    height: 1px;
    background: rgba(150,150,150,0.2);
}

/* ─── ラベル ────────────────────────────────── */
.stTextInput label,
.stSelectbox label,
.stDateInput label,
.stNumberInput label {
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    color: #9CA3AF !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}

/* ─── 入力フィールド ─────────────────────────── */
.stTextInput input,
.stNumberInput input,
.stDateInput input {
    border-radius: 10px !important;
    border: 1.5px solid #E5E7EB !important;
    font-size: 0.97rem !important;
    font-weight: 500 !important;
    padding: 0.65rem 0.9rem !important;
    background: #FAFAFA !important;
    color: #111827 !important;
    transition: all 0.15s ease !important;
    box-shadow: none !important;
}
.stTextInput input:focus,
.stNumberInput input:focus,
.stDateInput input:focus {
    border-color: #111827 !important;
    background: #fff !important;
    box-shadow: 0 0 0 3px rgba(17,24,39,0.06) !important;
}

/* ─── セレクトボックス ───────────────────────── */
.stSelectbox > div > div {
    border-radius: 10px !important;
    border: 1.5px solid #E5E7EB !important;
    font-size: 0.97rem !important;
    font-weight: 500 !important;
    background: #FAFAFA !important;
    color: #111827 !important;
}

/* ─── ボタン共通 ─────────────────────────────── */
.stButton > button {
    font-family: 'Inter', 'Noto Sans JP', sans-serif !important;
    font-size: 0.87rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em !important;
    border-radius: 10px !important;
    height: auto !important;
    padding: 0.75rem 1.2rem !important;
    transition: all 0.18s cubic-bezier(0.4, 0, 0.2, 1) !important;
    border: none !important;
    cursor: pointer !important;
    width: 100% !important;
}

/* 記録ボタン（primary） */
button[kind="primary"] {
    background: #111827 !important;
    color: #ffffff !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.08) !important;
    letter-spacing: 0.04em !important;
    font-size: 0.9rem !important;
    padding: 0.9rem 1.2rem !important;
}
button[kind="primary"]:hover {
    background: #1F2937 !important;
    box-shadow: 0 4px 16px rgba(0,0,0,0.18) !important;
    transform: translateY(-1px) !important;
}
button[kind="primary"]:active {
    transform: translateY(0) !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1) !important;
}

/* 追加ボタン・その他（secondary） */
button[kind="secondary"] {
    background: #ffffff !important;
    border: 1.5px solid #E5E7EB !important;
    color: #374151 !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04) !important;
}
button[kind="secondary"]:hover {
    border-color: #111827 !important;
    color: #111827 !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
    transform: translateY(-1px) !important;
}
button[kind="secondary"]:active {
    transform: translateY(0) !important;
}

/* 削除ボタン */
.del-btn .stButton > button {
    background: transparent !important;
    border: 1px solid #F3F4F6 !important;
    color: #D1D5DB !important;
    font-size: 0.78rem !important;
    padding: 0.5rem !important;
    box-shadow: none !important;
    letter-spacing: 0 !important;
}
.del-btn .stButton > button:hover {
    border-color: #FCA5A5 !important;
    color: #EF4444 !important;
    background: #FFF5F5 !important;
    transform: none !important;
    box-shadow: none !important;
}

/* ─── 合計ブロック ───────────────────────────── */
.total-block {
    background: #111827;
    border-radius: 14px;
    padding: 1.1rem 1.4rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin: 1rem 0 0.75rem;
}
.total-label-t {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #6B7280;
}
.total-amount-t {
    font-size: 1.5rem;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -0.02em;
}

/* ─── 完了画面 ──────────────────────────────── */
.done-wrap {
    padding: 2rem 0;
}
.done-title {
    font-size: 1.25rem;
    font-weight: 700;
    color: inherit;
    text-align: center;
    letter-spacing: -0.02em;
    margin-bottom: 0.3rem;
}
.done-sub {
    font-size: 0.85rem;
    color: #9CA3AF;
    text-align: center;
    margin-bottom: 1.6rem;
}
.done-list {
    background: #fff;
    border: 1px solid #E5E7EB;
    border-radius: 14px;
    overflow: hidden;
    margin-bottom: 1.4rem;
}
.done-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.85rem 1.1rem;
    border-bottom: 1px solid #F3F4F6;
}
.done-row:last-child { border-bottom: none; }
.done-cat { font-size: 0.9rem; color: #374151; font-weight: 500; }
.done-note { font-size: 0.75rem; color: #9CA3AF; margin-top: 0.1rem; }
.done-amt { font-size: 0.97rem; font-weight: 700; color: #111827; }

/* ─── 残高 島 ───────────────────────────────── */
.island {
    border: 1px solid #E5E7EB;
    border-radius: 16px;
    overflow: hidden;
    margin-bottom: 0.75rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.island-header {
    padding: 0.65rem 1.1rem;
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #9CA3AF;
    border-bottom: 1px solid #F3F4F6;
    background: rgba(0,0,0,0.02);
}
.island-row {
    display: flex;
    align-items: center;
    padding: 0.8rem 1.1rem;
    border-bottom: 1px solid #F3F4F6;
    gap: 0.75rem;
}
.island-row:last-child { border-bottom: none; }
.island-row-left { flex: 1; min-width: 0; }
.island-row-name {
    font-size: 0.82rem;
    font-weight: 500;
    margin-bottom: 0.35rem;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.island-bar-bg {
    height: 3px;
    background: rgba(150,150,150,0.15);
    border-radius: 999px;
    overflow: hidden;
}
.island-bar-fill {
    height: 100%;
    border-radius: 999px;
}
.island-bar-fill.safe    { background: #10B981; }
.island-bar-fill.warning { background: #F59E0B; }
.island-bar-fill.danger  { background: #EF4444; }
.island-row-right { text-align: right; flex-shrink: 0; }
.island-remaining {
    font-size: 1rem;
    font-weight: 700;
    letter-spacing: -0.02em;
    line-height: 1;
}
.island-remaining.safe    { color: #10B981; }
.island-remaining.warning { color: #D97706; }
.island-remaining.danger  { color: #EF4444; }
.island-used {
    font-size: 0.68rem;
    color: #9CA3AF;
    margin-top: 0.2rem;
}

/* ─── フッター ──────────────────────────────── */
.page-footer {
    text-align: center;
    padding: 1.5rem 0 0;
    font-size: 0.75rem;
    color: #D1D5DB;
}
.page-footer a {
    color: #9CA3AF;
    text-decoration: none;
    font-weight: 500;
}
.page-footer a:hover { color: #111827; }
</style>
""", unsafe_allow_html=True)

# ── セッション初期化 ────────────────────────────
if "entries" not in st.session_state:
    st.session_state.entries = [{"date": str(date.today()), "category": "食費", "amount": 0, "note": ""}]
if "done" not in st.session_state:
    st.session_state.done = None

# ── 完了画面 ───────────────────────────────────
if st.session_state.done:
    rows  = st.session_state.done
    total = sum(int(r[5]) for r in rows)

    rows_html = ""
    for r in rows:
        note_html = f'<div class="done-note">{r[2]}</div>' if r[2] else ""
        rows_html += f'''
        <div class="done-row">
          <div>
            <div class="done-cat">{r[3]}</div>
            {note_html}
          </div>
          <div class="done-amt">¥{int(r[5]):,}</div>
        </div>'''

    st.markdown(f'''
    <div class="done-wrap">
      <div class="done-title">記録しました</div>
      <div class="done-sub">{len(rows)}件　合計 ¥{total:,}</div>
      <div class="done-list">{rows_html}</div>
    </div>
    ''', unsafe_allow_html=True)

    if st.button("続けて入力する", use_container_width=True):
        st.session_state.done = None
        st.session_state.entries = [{"date": str(date.today()), "category": "食費", "amount": 0, "note": ""}]
        st.rerun()

    st.markdown(
        f'<div class="page-footer"><a href="https://docs.google.com/spreadsheets/d/{SPREAD_ID}/edit" target="_blank">スプレッドシートを開く →</a></div>',
        unsafe_allow_html=True
    )
    st.stop()

# ── ヘッダー ───────────────────────────────────
today = date.today()
st.markdown('<p class="page-header-label">Household Budget</p>', unsafe_allow_html=True)
st.title("支出を記録する")
st.markdown(f'<p class="page-header-date">{today.year}年{today.month}月{today.day}日（{DAYS_JA[today.weekday()]}）</p>', unsafe_allow_html=True)
st.divider()

# ── 残高ウィジェット ───────────────────────────
spending, spend_err  = get_monthly_spending(today.year, today.month)
if spend_err:
    st.warning(f"残高取得エラー: {spend_err}")

with st.expander("🔍 デバッグ（確認後に削除）", expanded=False):
    try:
        _gc = get_gc()
        _ws = _gc.open_by_key(SPREAD_ID).worksheet("明細")
        _rows = _ws.get_all_values()
        st.write("ヘッダー行:", _rows[0] if _rows else "なし")
        st.write("データ行（最初の3件）:", _rows[1:4] if len(_rows) > 1 else "なし")
        st.write("今月一致行数:", sum(1 for r in _rows[1:] if len(r)>0 and r[0].startswith("2026-05")))
    except Exception as ex:
        st.write("エラー:", ex)
islands_html = ""

for group in WATCH_GROUPS:
    islands_html += f'<div class="island"><div class="island-header">{group["label"]}</div>'
    for name in group["items"]:
        budget    = WATCH_BUDGETS[name]
        spent     = spending.get(name, 0)
        remaining = budget - spent
        pct       = min(spent / budget * 100, 100) if budget > 0 else 100
        status    = "danger" if remaining < 0 else ("warning" if pct >= 80 else "safe")
        sign      = "−" if remaining < 0 else ""
        label     = name.replace("お小遣い", "")
        islands_html += f"""
        <div class="island-row">
          <div class="island-row-left">
            <div class="island-row-name">{label}</div>
            <div class="island-bar-bg">
              <div class="island-bar-fill {status}" style="width:{pct:.1f}%"></div>
            </div>
          </div>
          <div class="island-row-right">
            <div class="island-remaining {status}">{sign}¥{abs(remaining):,}</div>
            <div class="island-used">/ ¥{budget:,}</div>
          </div>
        </div>"""
    islands_html += "</div>"

st.markdown(islands_html, unsafe_allow_html=True)

# ── エントリ入力 ───────────────────────────────
delete_idx = None

for i, e in enumerate(st.session_state.entries):
    st.markdown(f'<div class="entry-card"><div class="card-index">Item {i+1:02d}</div></div>', unsafe_allow_html=True)

    d_val = date.fromisoformat(e["date"]) if e["date"] else date.today()
    st.session_state.entries[i]["date"] = str(
        st.date_input("日付", value=d_val, key=f"date_{i}")
    )

    cat_idx = CATEGORIES.index(e["category"]) if e["category"] in CATEGORIES else 0
    st.session_state.entries[i]["category"] = st.selectbox(
        "費目", CATEGORIES, index=cat_idx, key=f"cat_{i}"
    )

    st.session_state.entries[i]["amount"] = st.number_input(
        "金額（円）", min_value=0, value=e["amount"], step=10, key=f"amt_{i}"
    )

    st.session_state.entries[i]["note"] = st.text_input(
        "補足", value=e["note"], placeholder="店舗名・メモなど", key=f"note_{i}"
    )

    if i > 0:
        st.markdown('<div class="del-btn">', unsafe_allow_html=True)
        if st.button("削除", key=f"del_{i}", use_container_width=True):
            delete_idx = i
        st.markdown('</div>', unsafe_allow_html=True)

if delete_idx is not None:
    st.session_state.entries.pop(delete_idx)
    st.rerun()

# ── 追加ボタン ─────────────────────────────────
st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
if st.button("＋  項目を追加", use_container_width=True):
    st.session_state.entries.append({"date": str(date.today()), "category": "食費", "amount": 0, "note": ""})
    st.rerun()

# ── 合計 ───────────────────────────────────────
valid = [e for e in st.session_state.entries if e["amount"] > 0]
if valid:
    total = sum(e["amount"] for e in valid)
    st.markdown(f'''
    <div class="total-block">
      <span class="total-label-t">合計</span>
      <span class="total-amount-t">¥{total:,}</span>
    </div>
    ''', unsafe_allow_html=True)
else:
    st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)

# ── 記録ボタン ─────────────────────────────────
if st.button("記録する", type="primary", use_container_width=True):
    if not valid:
        st.error("金額を入力してください")
    else:
        with st.spinner(""):
            try:
                rows = append_to_sheet(valid)
                st.session_state.done = rows
                st.session_state.entries = [{"date": str(date.today()), "category": "食費", "amount": 0, "note": ""}]
                get_gc.clear()
                get_monthly_spending.clear()   # 残高を即時更新
                st.rerun()
            except Exception as ex:
                st.error(f"エラー: {ex}")

# ── フッター ───────────────────────────────────
st.markdown(
    f'<div class="page-footer"><a href="https://docs.google.com/spreadsheets/d/{SPREAD_ID}/edit" target="_blank">スプレッドシートを開く →</a></div>',
    unsafe_allow_html=True
)
