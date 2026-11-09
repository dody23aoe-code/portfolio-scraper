# 書籍スクレイピングのPythonサンプル

[Books to Scrape](https://books.toscrape.com/) の全カテゴリ共通の一覧を最後のページまで取得し、各書籍の詳細から題名・価格・在庫・評価・カテゴリを読み取ります。カテゴリ別ページを重複して回る必要はありません。価格と評価はサイトの練習用データです。

初回は全書籍をCSVとExcelに保存します。2回目以降は前回の正常終了時と比較し、追加・更新・削除された行だけを保存します。ネットワーク取得は毎回全件行います。

## 導入

Python 3.11以上を用意し、このリポジトリのルートで実行します。npm / pnpm は不要です。

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -v
```

macOS / Linux:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m unittest discover -v
```

## 実行例

```powershell
# robots.txtだけを確認する（書籍は取得しない）
.\.venv\Scripts\python.exe scraper.py --check-robots

# 初回: 全件をCSV・Excelに保存
.\.venv\Scripts\python.exe scraper.py

# 2回目以降: 前回との差分だけをCSV・Excelに保存
.\.venv\Scripts\python.exe scraper.py --config config.json
```

macOS / LinuxではPythonの実行部分を `.venv/bin/python` に置き換えてください。正常終了は終了コード0、取得・解析・保存の失敗は1です。

現在の一覧規模（1,000冊・50ページ）では約1,051回のHTTPリクエストが必要です。設定の待機時間だけでも約20分かかり、通信時間や再試行でさらに長くなります。進捗は一覧ページごとに表示されます。

## 設定

`config.json` を編集します。

| 項目 | 既定値 | 内容 |
| --- | --- | --- |
| `base_url` | `https://books.toscrape.com/` | この練習サイトに限定 |
| `user_agent` | `BooksPracticeSample/1.0` | robots判定にも使用 |
| `request_interval_seconds` | `1.1` | 前リクエスト終了から次の開始までの最低秒数。1未満は拒否 |
| `timeout_seconds` | `30` | requestsの接続・読み取りタイムアウト秒数 |
| `max_retries` | `3` | 初回に加える再試行数。0〜10 |
| `backoff_seconds` | `2` | 再試行待機の基準。2、4、8秒と増加 |
| `output_dir` | `output` | スクリプトのあるフォルダ内の出力サブフォルダ |

## 取得時の規則

- 最初にrobots.txtを取得し、すべての一覧・詳細URLを `urllib.robotparser` で判定します。取得禁止URLがあれば、その実行全体を中止します。
- robots.txtが404または410の場合は規則が存在しないものとして扱います。403・通信失敗・再試行後のサーバーエラーなどでは中止します。
- `Crawl-delay` と `Request-rate` が解析された場合は、設定より長い待機時間を採用します。robots.txtの取得と再試行にも1秒以上の間隔を適用します。並列取得はしません。
- 通信エラー、408、429、5xxは上限つきで再試行します。`Retry-After` の秒数またはHTTP日時も考慮します。その他のHTTPエラーは即時中止します。
- 自動リダイレクトは無効です。3xxは中止し、別ホストにはアクセスしません。
- 必須項目の欠落、ページ送りの循環、総件数の変動・取得数の不一致はエラーとして中止します。失敗した取得を「削除」として差分保存しません。

## 出力例

```text
output/
  latest.json
  runs/
    <UTC実行日時と識別子>/
      changes.csv
      changes.xlsx
```

初回のCSV例（UTF-8 BOM付き）:

```csv
change,url,title,price_gbp,stock,rating,category
added,https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html,A Light in the Attic,51.77,In stock (22 available),3,Poetry
```

Excelにも同じ列を保存します。`price_gbp` は小数2桁を保持する文字列（通貨GBP）、`rating` は1〜5の整数です。在庫は詳細ページの文字列を空白整理して保存します。CSVで数式として解釈されうる文字列には先頭に `'` を付け、Excelでは文字列型で保存します。比較用JSONには元の値を保持します。

書籍URLを比較キーにし、5項目のいずれかが変われば `updated` として変更後の行を保存します。新規は `added`、消えた書籍は `removed` として前回の行を保存します。URLが変わると削除と追加になります。変化がなければCSV・Excelともヘッダーのみです。

`latest.json` は比較に必要な全件データと最後に成功した実行IDを保持します。CSV・Excelの作成後に置換して更新します。履歴CSV・Excelは差分のみで、`latest.json` は最新の全件スナップショットです。

保存途中に異常終了した場合、未完成の実行フォルダや一時ファイルが残ることがあります。`latest.json` の実行IDだけが最新の確定結果です。同時起動は `output/.lock` で防止します。強制終了でロックが残った場合は、別の実行がないことを確認してロックファイルを削除してください。取得の途中再開は行わず、再実行時に先頭から取得します。

## オフラインテスト

```powershell
.\.venv\Scripts\python.exe -m unittest discover -v
```

実サイトから保存したHTMLと最小HTMLを `tests/fixtures/` に置いています。実HTTP通信を禁止し、仮想時計で待機を検証するため、ネット接続や実際の待機は不要です。テスト用の一時ファイルはプロジェクト内の `.tmp/` に作成・削除します。

検証対象はHTML解析、全ページ走査、重複排除、robots判定、リクエスト間隔、再試行上限、差分判定、CSV・Excel、保存失敗時の比較基準維持です。

## 公開するファイル

本体・設定・依存関係・README・テストと保存HTMLが公開対象です。取得結果、仮想環境、キャッシュ、`REPORT.md` は `.gitignore` で除外しています。GitHubリポジトリの作成・pushを行う処理は含みません。
