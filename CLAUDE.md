# 家計管理システム ルールブック

## プロジェクト概要
レシート写真 → Claude が費目判定 → Google スプレッドシートに自動記録する家計管理システム。

- **アプリ**: https://kakeibo-yasmeshhjqcldacpmzftfq.streamlit.app
- **スプレッドシート**: https://docs.google.com/spreadsheets/d/1hbn526buIkrM2VUg10R2V8X6ifW1vCzBOGlNUWSv-3Q/edit
- **GitHub**: https://github.com/kentaro-kage/kakeibo

---

## ⚠️ 最重要：明細シートの列構成（7列）

| A: 日付 | B: 曜日 | C: 店舗名 | D: 品目 | E: 費目 | F: 金額(税込) | G: メモ |

- **費目 = 列E（index 4）**、**金額 = 列F（index 5）**
- app.py の読み取り: `r[4]` = 費目、`r[5]` = 金額
- SUMPRODUCT 数式: `明細!$E$2:$E$2000`（費目）、`明細!$F$2:$F$2000`（金額）

---

## ファイル構成
```
kakeibo/
├── app.py              ← Streamlit アプリ本体（メイン）
├── rebuild.py          ← スプレッドシート再構築（⚠️明細データが消える）
├── restore_meisai.py   ← backups/ から明細を復元
├── config.json         ← スプレッドシートID・カテゴリ
├── token.json          ← Google OAuth トークン（.gitignore済み・機密）
├── credentials.json    ← OAuth クライアント情報（.gitignore済み・機密）
└── backups/            ← 明細の自動バックアップCSV
```

---

## 費目カテゴリ（20種）
家賃 / 光熱費 / 通信費 / 保育料 / 保険 / 外貨積立 / 税金 / 医療費 / 衣料美容費 / 食費 / 日用品 / 健太郎お小遣い / 藍子お小遣い / 陽和お小遣い / サブスク代 / 外食費 / 自動車 / ペット / 自己投資 / その他

---

## 月次予算（残金ウィジェット対象）
| 費目 | 予算 |
|------|------|
| 食費 | ¥30,000 |
| 外食費 | ¥15,000 |
| 健太郎お小遣い | ¥10,000 |
| 藍子お小遣い | ¥10,000 |
| 陽和お小遣い | ¥10,000 |

---

## Streamlit Cloud 設定

Secrets（`[google_token]`セクション）に token.json の内容を設定。  
**token は数時間で期限切れになる。**  
ローカルで何かスクリプトを実行すると token.json が自動更新されるので、その値を Streamlit Cloud の Secrets に反映する。

---

## デプロイ方法
```bash
git add app.py
git commit -m "変更内容"
git push "https://kentaro-kage:$(gh auth token)@github.com/kentaro-kage/kakeibo.git" main
# → 1〜2分で自動デプロイ
```

---

## よくあるトラブル対処

| 症状 | 原因 | 対処 |
|------|------|------|
| アプリの残金が変わらない | Streamlit Cloud の token 期限切れ | Secrets の expiry・token を更新 |
| 他タブ（予算プランなど）に反映されない | SUMPRODUCT の列参照がずれている | rebuild.py でE/F参照に修正して数式だけ更新 |
| 明細データが消えた | rebuild.py を実行した | `python3 restore_meisai.py` で復元 |
| Googleテーブル形式エラー | 明細シートが「表_1」形式 | API で deleteTable してから書き込む |
| 金額欄の値が消える | `value=None` を固定で渡している | `value=e["amount"]` で入力値を保持 |

---

## 引き継ぎ用詳細ドキュメント
→ `ONBOARDING.md` を参照
