# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## このリポジトリについて

RoomRadar（日大理工学部 空き教室検索）— 日本大学理工学部向けの、リアルタイム空き教室検索サービス。
曜日・時限・校舎を選ぶと授業が入っていない教室を表示し、その上に簡易的な「仮予約」「実は使われていた報告」
機能を載せている。

このリポジトリには、ブランドは共有しているがコードは共有していない、3つの独立したデプロイ対象が含まれる。

1. **`app.py`** — 実際のFlaskアプリケーション本体（検索＋予約/報告API）。Renderにデプロイされている
   （`https://nust-room-search.onrender.com`）。
2. **`index.html`** — GitHub Pagesでホストされている静的なランディングページ
   （`csko24143-droid.github.io/nust-room-search`）。実際の検索機能はRender上のアプリへリンクする形。
   日本語/英語/中国語の3言語対応（CSSクラスで切り替え。詳細は後述）。
3. **`dashboard.html`** — `noindex`指定の静的アナリティクスダッシュボード（Chart.js使用）。`data/`配下の
   JSONを読み込んでInstagram/GAのインサイトを表示する。メインUIからはリンクされていない。

これら3つの間でビルドシステムは共有されていない。それぞれ単体で完結したHTMLファイル
（`app.py`の場合はHTML/CSS/JSをPythonの文字列テンプレートとして埋め込んだPythonファイル）。

## コマンド

このリポジトリにテストスイート・linter・ビルドステップは存在しない。

```bash
pip install -r requirements.txt   # flask, pandas, openpyxl, gunicorn

python app.py                     # ローカルでFlaskアプリを起動（PORT環境変数を読む、デフォルト10000）
gunicorn app:app                  # 本番用エントリポイント（Renderが使用）
```

Instagram連携スクリプト（手動実行、または`.github/workflows/`内のGitHub Actionsから実行）には
`IG_ACCESS_TOKEN`と`IG_ACCOUNT_ID`の環境変数が必要。

```bash
python scripts/fetch_instagram_posts.py
python scripts/fetch_instagram_insights.py
python scripts/post_to_instagram.py --image-url <url> --caption "<text>"
```

## アーキテクチャ: `app.py`

単一ファイルのFlaskアプリ。ルーティング、DBアクセス、ページ本体
（`HTML_TEMPLATE`という、`render_template_string`で描画されるJinja文字列テンプレート）が
すべてこの1ファイルに収まっている。`templates/`や`static/`ディレクトリは存在しない。

**3つの独立したSQLiteデータベース、それぞれライフサイクルが異なる：**

- `schedule_final.db` — gitにコミットされている、読み取り専用の正データ。テーブルは2つ：
  `schedules`（時間割。カラム名は日本語：学科/履修期名/曜日/時限/教室/校舎/科目名）と
  `classrooms`（`name`, `building`）。`classroom_data.xlsx` / `summry_classrooms.xlsx`はこのDBの
  元になった生の表データだが、実行時にコードから読まれることはない。時間割を更新したい場合は
  `schedule_final.db`自体を直接作り直す/置き換える必要がある。
- `reservations.db` / `reports.db` — `init_reserve_db()` / `init_reports_db()`によって実行時に
  作成される、`.gitignore`対象のDB。「仮予約」と「実は使われていた報告」機能のバックエンド。
  どちらも`cancel_code`方式：作成時にランダムな6文字コードをクライアントへ返し`localStorage`に
  保存、後で削除/取り消しする際に同じコードが必要。認証は無く、コードを知っていれば誰でも
  取り消せる仕組み。

**学期・時間のロジック：** `ACTIVE_TERMS`と`get_active_terms()`が、前期/後期どちらの時間割行を
適用するかを、ハードコードされた月日の範囲（4/1〜9/20＝前期）で判定する。`PERIODS`は時限番号
（1〜6）をJSTの開始/終了時刻にマッピングし、`period_end_dt()`がこれを使って予約/報告がいつ自動
失効するかを計算する（`cleanup_expired()` / `cleanup_reports()`が該当リクエストの先頭で呼ばれる
形でクリーンアップされる。バックグラウンドジョブは無い）。

**レート制限とバリデーションはプロセス内・メモリ上**（`_rate_store`、IPをキーにした
`collections.defaultdict(list)`）。これは単一ワーカープロセスでのみ正しく機能する仕組み。
将来Render/gunicornをマルチワーカーにスケールした場合、レート制限はワーカーごとに別々になり
グローバルには効かなくなる。`VALID_DAYS`/`VALID_PERIODS`/`VALID_BUILDINGS`が全リクエスト
パラメータのホワイトリストになっている。

**APIルート**（`/api/reserve`, `/api/reserve/cancel`, `/api/reserve/list`, `/api/report`,
`/api/report/cancel`）は、`HTML_TEMPLATE`内のインライン`<script>`から呼ばれる素のJSON
POST/GETエンドポイント。`/`（GET/POST）が検索ページ本体で、POSTは検索フォームの送信に対応する。

## アーキテクチャ: Instagram連携（`scripts/` + `.github/workflows/`）

3つのスクリプトがInstagram Graph APIをラップしており、それぞれ`graph.instagram.com`を試した後
`graph.facebook.com`にフォールバックする（`HOSTS`リストのパターンが3ファイルすべてに重複している）。

- `fetch_instagram_posts.py` — デバッグ/確認用スクリプト。最近の投稿を標準出力に表示するだけ。
  手動の`workflow_dispatch`のみ。
- `fetch_instagram_insights.py` — 本体のデータパイプライン：`data/instagram_history.json`に
  日次スナップショットを追記し、`data/instagram_posts.json`を最新の投稿パフォーマンスで上書きする。
  `instagram-insights.yml`によりcron実行（月・木 09:00 JST）され、その後更新されたJSONを
  `github-actions[bot]`としてブランチへ直接コミットする。`dashboard.html`はこの2つのJSONを
  クライアント側で読み込んでグラフを描画する。
- `post_to_instagram.py` — 画像URL＋キャプションを指定してフィード投稿を公開する。
  `image_url`/`caption`を入力とする手動の`workflow_dispatch`のみで、自動実行されることは無い。

`data/instagram_history.json`または`data/instagram_posts.json`の形式を変更する場合は、
書き込み側（`fetch_instagram_insights.py`）と読み込み側（`dashboard.html`のJS）を両方
同時に更新すること。

## 規約

- DB/テーブル/カラム識別子、およびサーバー側のエラー/ログ文字列の多くは日本語になっている。
  新しいフィールドを追加する際も英語表記に置き換えるのではなく、既存の命名に合わせること。
- 日時の扱いはすべて`JST`（`datetime.timezone(timedelta(hours=9))`）を経由する。ユーザーに
  見える処理でnaiveな`datetime.now()`を使わないこと。
- `index.html`の言語切り替えは、`<body>`に`lang-en`/`lang-zh`クラスを付与し、CSS
  （ブロック要素は`.ja`/`.en`/`.zh`、インライン要素は`span.ja`/`span.en`/`span.zh`）で
  事前に翻訳済みのテキストを出し分ける仕組みで、i18nライブラリや実行時の文字列ルックアップは
  使っていない。新しい文言を追加する場合は、既存パターンに従って3言語分をインラインで
  すべて追加する必要がある。
