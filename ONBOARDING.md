# 家計管理システム 引き継ぎガイド

## システム概要

レシートを撮影して Claude に送るだけで、費目別に自動分類・Google スプレッドシートへ記録する家計管理システム。

- **アプリURL**: https://kakeibo-yasmeshhjqcldacpmzftfq.streamlit.app
- **スプレッドシート**: https://docs.google.com/spreadsheets/d/1hbn526buIkrM2VUg10R2V8X6ifW1vCzBOGlNUWSv-3Q/edit
- **GitHubリポジトリ**: https://github.com/kentaro-kage/kakeibo

---

## ディレクトリ構成

```
/Users/kentarokage/Private/kakeibo/
├── app.py              ← Streamlit アプリ本体
├── rebuild.py          ← スプレッドシート全シート再構築スクリプト
├── restore_meisai.py   ← 明細データ復元スクリプト
├── config.json         ← スプレッドシートID・カテゴリ設定
├── token.json          ← Google OAuth2 トークン（機密）
├── credentials.json    ← Google OAuth2 クライアント情報（機密）
├── requirements.txt    ← 依存パッケージ
└── backups/            ← 明細シートの自動バックアップCSV
```

---

## スプレッドシート構成

### 明細シート（データ入力先）
**7列構成**（重要：6列ではなく7列）

| A: 日付 | B: 曜日 | C: 店舗名 | D: 品目 | E: 費目 | F: 金額(税込) | G: メモ |
|---------|---------|---------|---------|---------|------------|---------|

- **費目は E 列（index 4）**、金額は F 列（index 5）
- SUMPRODUCT 数式は `明細!$E$2:$E$2000`（費目）、`明細!$F$2:$F$2000`（金額）を参照

### その他シート
- **ダッシュボード** - 今月の収支サマリー
- **収支管理** - 月次収支グラフ・実績
- **予算プラン** - 費目別予算 vs 実績
- **月別集計** - 5〜12月の費目別集計

---

## 費目カテゴリ（20種）

```
家賃 / 光熱費 / 通信費 / 保育料 / 保険 / 外貨積立 / 税金
医療費 / 衣料美容費 / 食費 / 日用品 / 健太郎お小遣い
藍子お小遣い / 陽和お小遣い / サブスク代 / 外食費 / 自動車
ペット / 自己投資 / その他
```

---

## アプリ（app.py）の重要仕様

### 残金ウィジェット（WATCH_BUDGETS）
```python
WATCH_BUDGETS = {
    "食費":           30000,
    "外食費":         15000,
    "健太郎お小遣い": 10000,
    "藍子お小遣い":   10000,
    "陽和お小遣い":   10000,
}
```
「食費島」「お小遣い島」の2グループで表示。

### Google 認証
- ローカル実行: `token.json` を直接読み込み
- Streamlit Cloud: `st.secrets["google_token"]` から読み込み
- トークン期限切れ時は自動更新（`creds.refresh(Request())`）

### シートへの書き込み（append_to_sheet）
```python
# 7列で書き込む: [日付, 曜日, 店舗名, 品目, 費目, 金額, メモ]
rows.append([date, day_ja, note, category, category, amount, ""])
```

### 残金読み込み（get_monthly_spending）
```python
cat = r[4]   # 費目（E列）
amt = r[5]   # 金額（F列）
```

---

## Streamlit Cloud 設定

### Secrets（TOML形式）
```toml
[google_token]
token         = "ya29.xxx..."
refresh_token = "1//xxx..."
token_uri     = "https://oauth2.googleapis.com/token"
client_id     = "700254398924-xxx.apps.googleusercontent.com"
client_secret = "GOCSPX-xxx"
scopes        = ["https://www.googleapis.com/auth/spreadsheets"]
universe_domain = "googleapis.com"
account       = ""
expiry        = "2026-xx-xxTxx:xx:xx.xxxxZ"
```

**token は期限切れになる。token.json の内容を定期的に更新すること。**

---

## よくある問題と対処

| 問題 | 原因 | 対処 |
|------|------|------|
| 残金が変わらない | キャッシュ（TTL=60秒）or トークン期限切れ | アプリ再起動 or Secrets のtoken更新 |
| #VALUE! エラー | SUMPRODUCT が列D/Eを参照（旧仕様） | rebuild.py で E/F に修正済み |
| 明細データが消える | rebuild.py 実行時に自動クリア | restore_meisai.py で backups/ から復元 |
| テーブル形式エラー | Google Sheets の「表_1」が列型制約を強制 | API で deleteTable してから書き込む |

---

## 運用方法

1. レシート撮影 → Claude にチャットで送信
2. Claude が費目を判定してアプリ経由でシートに記録
3. アプリURL: https://kakeibo-yasmeshhjqcldacpmzftfq.streamlit.app

### rebuild.py を実行する場合の注意
- 実行前に backups/ に自動バックアップが取られる
- 明細データが一度クリアされる → 実行後に restore_meisai.py で復元
- `setup_monthly` / `setup_budget_plan` / `setup_shuushi` / `setup_dashboard` のみ個別実行推奨

---

## デプロイ手順

```bash
# コード変更後
git add app.py
git commit -m "変更内容"
git push "https://kentaro-kage:$(gh auth token)@github.com/kentaro-kage/kakeibo.git" main
# → Streamlit Cloud が自動デプロイ（約1〜2分）
```
