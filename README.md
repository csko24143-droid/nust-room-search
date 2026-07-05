# RoomRadar — 日大理工 空き教室検索

日本大学理工学部の空き教室を検索できるWebサービス「RoomRadar」のリポジトリです。
検索アプリ本体（Flask）と、広報用のランディングページ・分析ダッシュボード（静的HTML）、
Instagram/GA4の自動化スクリプトを含みます。

## 構成

```
├── app.py                    # 検索アプリ本体（Flask）。教室検索・利用予約・報告API
├── schedule_final.db         # 授業スケジュールDB（SQLite・読み取り専用の元データ）
├── index.html                # ランディングページ（静的・GitHub Pages等で公開）
├── dashboard.html            # 分析ダッシュボード（data/ のJSONを読み込んで可視化）
├── rr_logo.png               # ロゴ（index.html から参照）
├── classroom_data.xlsx       # 元データ（教室データ・コードからは未参照）
├── summary_classrooms.xlsx   # 元データ（教室サマリ・コードからは未参照）
├── data/                     # ワークフローが自動更新する分析データ（JSON）
│   ├── ga4_history.json          # GA4 日次推移
│   ├── ga4_breakdown.json        # GA4 内訳（イベント・検索条件・流入元）
│   ├── instagram_history.json    # Instagram インサイト日次スナップショット
│   └── instagram_posts.json      # Instagram 最新投稿パフォーマンス
├── scripts/                  # 自動化スクリプト（各ファイル冒頭のdocstring参照）
│   ├── fetch_ga4_metrics.py       # GA4データ取得 → data/ へ保存
│   ├── fetch_instagram_insights.py# IGインサイト取得 → data/ へ保存
│   ├── fetch_instagram_posts.py   # IG投稿一覧の取得・表示
│   ├── post_to_instagram.py       # IGへフィード/ストーリーズ投稿
│   └── make_story_promo.py        # 新着投稿の告知ストーリーズ画像を生成
├── assets/                   # 画像素材（QR・ポスター・ストーリーズ生成物）
└── .github/workflows/        # 定期実行・手動実行のGitHub Actions
```

## 検索アプリ（app.py）

- 曜日・時限・棟を指定して空き教室を検索
- 「この教室を使う」予約機能（`reservations.db` を実行時に自動生成）
- 「実際は使えなかった」報告機能（`reports.db` を実行時に自動生成、一定数でグレーアウト）
- レート制限・入力バリデーションつき

### ローカル実行

```bash
pip install -r requirements.txt
python app.py            # http://localhost:10000
```

本番は gunicorn を想定（`gunicorn app:app`）。ポートは `PORT` 環境変数で指定（デフォルト10000）。

## GitHub Actions ワークフロー

| ファイル | トリガー | 内容 |
|---|---|---|
| `ga4-fetch.yml` | 月・木 JST 9:10 / 手動 | GA4データを取得し `data/` にコミット |
| `instagram-insights.yml` | 月・木 JST 9:00 / 手動 | IGインサイトを取得し `data/` にコミット |
| `instagram-fetch.yml` | 手動 | IG投稿一覧をログに表示（確認用） |
| `instagram-post.yml` | 手動 | IGへ投稿。フィード投稿時は告知ストーリーズも自動生成・投稿 |

### 必要なSecrets

| Secret | 用途 |
|---|---|
| `IG_ACCESS_TOKEN` | Instagram Graph API アクセストークン |
| `IG_ACCOUNT_ID` | InstagramビジネスアカウントID |
| `GA4_PROPERTY_ID` | GA4プロパティID |
| `GA4_SA_KEY` | GA4閲覧者権限のサービスアカウントJSON鍵（中身そのもの） |

## 備考

- Instagram投稿には自主創造プロジェクトの規定により `#日大生プロジェクト` タグが自動付与されます（`scripts/post_to_instagram.py`）
- `dashboard.html` は `./data/*.json` を相対パスで読むため、リポジトリルートから配信すること
