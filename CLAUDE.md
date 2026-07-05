# CLAUDE.md

RoomRadar（日大理工 空き教室検索）のリポジトリ。全体像は README.md を参照。

## 重要な前提

- `app.py` は単一ファイルのFlaskアプリ。HTMLは `render_template_string` でインライン定義されており、`index.html` とは別物（index.htmlは広報用ランディングページ）。
- `schedule_final.db` は読み取り専用の元データ。`reservations.db` / `reports.db` は実行時に自動生成される（gitignore済み）。触らないこと。
- `data/` 配下のJSONはGitHub Actionsが自動コミットで更新する。手動編集しない。
- `index.html` / `dashboard.html` はリポジトリルートからの相対パス（`rr_logo.png`, `./data/*.json`）に依存。ファイル移動時は参照も更新すること。

## 実行・確認

```bash
pip install -r requirements.txt
python app.py   # http://localhost:10000 （PORT環境変数で変更可）
```

テストスイート・リンタ設定は現状なし。動作確認は起動＋主要ルート（`/`, `/api/reserve/list`）の応答で行う。

## 規約

- コミットメッセージは日本語（既存ログの体裁に合わせる: `chore: ...`, `fix: ...` など）
- Instagram投稿スクリプトの `#日大生プロジェクト` タグは大学規定による必須タグ。削除しない。
