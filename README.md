# 有給管理ツール / Paid Leave Manager

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-ES5-yellow.svg)](https://developers.google.com/apps-script)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Googleスプレッドシートで動作する有給休暇管理ツールです。日付と増減数を入力するだけで、自動的に残日数計算、付与日別管理、失効判定を行います。

A Google Sheets-based paid leave management tool that automatically calculates remaining days, manages grant dates, and tracks expiration.

## 機能 / Features

- ✅ **自動計算** - 入力時に自動的に残日数を計算
- ✅ **付与日別管理** - 各付与日ごとの残日数と使用状況を追跡
- ✅ **FIFO消費** - 古い付与日から順に自動的に消費
- ✅ **失効管理** - 2年後の失効日を自動計算
- ✅ **ステータス表示** - 有効/期限間近/使用済/失効済を自動判定
- ✅ **カスタムメニュー** - 手動再計算機能

## スクリーンショット / Screenshot

```
|   A    |    B    |    C    |  D  |   |    F    |    G    |   H   |   I   |     J      |    K     |
|--------|---------|---------|-----|---|---------|---------|-------|-------|------------|----------|
| 日付   | 増減数  | 残日数  | メモ|   | 付与日  |付与日数 | 使用済| 残日数| 失効予定日 | ステータス|
|--------|---------|---------|-----|---|---------|---------|-------|-------|------------|----------|
|2024/1/5|   10    |   10    |     |   |2024/1/5 |    10   |   3   |   7   | 2026/1/5   |   有効   |
|2024/3/10|   -3    |    7    |休暇 |   |         |         |       |       |            |          |
```

※ F-K列は「付与日別管理」で、A-D列の入力に基づいて自動生成されます。上記例では2024/1/5に10日付与後、2024/3/10に3日使用したため、付与日2024/1/5の使用済=3、残日数=7となります。

## セットアップ / Setup

### 1. スプレッドシートの作成 / Create Spreadsheet

1. [Googleスプレッドシート](https://sheets.google.com)を開く
2. 「空白」から新規スプレッドシートを作成
3. シート名を「有給管理」に変更（任意）
4. 1行目に以下のヘッダーを入力:

| 列 | ヘッダー | 説明 |
|----|---------|------|
| A | 日付 | 付与日または利用日 |
| B | 増減数 | 正の数=付与、負の数=利用 |
| C | 残日数 | 自動計算（入力不要） |
| D | メモ | 任意の備考 |
| E | （空白） | 仕切り列 |
| F | 付与日 | 自動生成（入力不要） |
| G | 付与日数 | 自動生成（入力不要） |
| H | 使用済 | 自動生成（入力不要） |
| I | 残日数 | 自動生成（入力不要） |
| J | 失効予定日 | 自動生成（入力不要） |
| K | ステータス | 自動生成（入力不要） |

### 2. Apps Scriptの設定 / Setup Apps Script

1. メニューから **拡張機能** → **Apps Script** を選択
2. `appsscript.js` の内容をエディタにコピー＆ペースト
3. **保存**（⌘/Ctrl + S）をクリック
4. プロジェクト名を変更（例: "有給管理"）

### 3. 権限の付与 / Authorize

1. Apps Scriptエディタで **実行** → `onOpen` を選択して実行
2. 承認画面が表示されたら、「Review Permissions」をクリック
3. Googleアカウントを選択し、権限を許可

## 使い方 / Usage

**有給の付与 / Grant Paid Leave:**
- A列に付与日（例: `2024/01/05`）を入力
- B列に付与日数（例: `10`）を入力
- C列とF-K列が自動的に計算・更新される

**有給の利用 / Use Paid Leave:**
- A列に利用日（例: `2024/03/10`）を入力
- B列にマイナスの利用日数（例: `-3`）を入力
- 古い付与日から自動的に消費される

**メニュー / Menu:**
- スプレッドシート上部に「有給管理」メニューが表示
- **有給管理** → **再計算** で全データを手動再計算

## システム構成 / Architecture

```
onOpen()               # カスタムメニューを追加

onEdit(e)              # 編集イベントを検知
    ↓
autoProcess()          # メイン処理関数
    ↓
├── readTransactions()        # A-D列からデータ読み込み
├── calculateRunningTotals()  # 残日数を計算
├── writeRunningTotals()      # C列に書き込み
├── processGrants()           # FIFO消費処理
├── buildGrantSummaries()     # サマリーデータを構築
└── writeGrantSummaries()     # F-K列に書き込み
```

## ステータス判定ロジック / Status Logic

| 条件 | ステータス |
|------|-----------|
| 残日数 ≤ 0 | 使用済 |
| 失効日 ≤ 今日 | 失効済 |
| 失効日 ≤ 今日 + 30日 | 期限間近 |
| その他 | 有効 |

## ファイル構成 / File Structure

```
holiday/
├── appsscript.js    # メインコード / Main code
├── README.md        # このファイル / This file
├── DESIGN.md        # 設計ドキュメント / Design document
├── AGENTS.md        # 開発ガイドライン / Development guidelines
├── LICENSE          # MITライセンス / MIT License
└── .gitignore       # Git除外設定 / Git ignore
```

## 技術仕様 / Technical Details

- **言語**: JavaScript (ES5 - Google Apps Script準拠)
- **インデックス**: 1-based（Google Sheets準拠）
- **消費方式**: FIFO（先入れ先出し法）
- **失効期間**: 付与日から2年
- **警告期間**: 失効日30日前

## 注意事項 / Notes

- 日付は日付形式（YYYY/MM/DD）で入力してください（文字列ではなく）
- 増減数は数値で入力してください
- 正の数 = 有給付与、負の数 = 有給利用
- 有給は付与日から2年で失効
- 利用時は古い付与日から順に消費（FIFO方式）
- ヘッダー行（1行目）は自動的に無視されます

## ライセンス / License

MIT License - 詳細は [LICENSE](LICENSE) を参照してください。

## 貢献 / Contributing

IssueやPull Requestは歓迎します！

バグ報告や機能追加のご要望は [Issues](../../issues) にてお知らせください。

---

**開発者向け情報**: 詳細な開発ガイドラインは [AGENTS.md](AGENTS.md) を参照してください。
