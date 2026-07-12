# RoomRadar — プロジェクト指示

日本大学理工学部の空き教室検索サービス。学生開発・**テスト運用中の非公式サービス**（自主創造プロジェクトの一環）。ユーザーとのやり取りは日本語で行う。

## 構成と配信先

- `app.py` — Flask検索アプリ本体。**Render**（https://nust-room-search.onrender.com ）でホスティング。mainへのpushで自動デプロイ。
- `index.html` / `dashboard.html` — LPと運営用アナリティクス。**GitHub Pages**（https://csko24143-droid.github.io/nust-room-search/ ）で配信。
- `schedule_final.db` — 時間割DB（検索の元データ）。`data/source/*.xlsx` が原本。
- `reservations.db` / `reports.db` — 実行時に自動生成される揮発データ。**コミットしない**（.gitignore済み）。
- `data/*.json` — Instagram/GA4の集計。GitHub Actionsが自動更新するため、手元と競合したら基本的に新しい方（項目が多い方）を採用。

## 守るべきルール

1. **「テスト運用中・非公式」表記を消さない**: app.py・index.html・dashboard.html・ポスターに常設。学生課/教務課と協議中のため誤認防止が必須。
2. **Instagram投稿キャプションには `#日大生プロジェクト` が必須**（自主創造プロジェクトの規定）。`scripts/post_to_instagram.py` が自動付与するので、この機構を壊さない。ストーリーズにはキャプション自体が無いので対象外。
3. **公開文言では「予約」と言い切らない**: アプリの仮予約は非公式であり教室の使用権を保証しない旨を必ず併記。
4. コミットは `git config user.email noreply@anthropic.com && git config user.name Claude` で行う（Verified表示のため）。

## Instagram運用

- 投稿は `.github/workflows/instagram-post.yml` を workflow_dispatch で実行（`post_type`: feed/story、`also_story`: フィード投稿と同時に告知ストーリーを自動生成・投稿）。
- 画像URLはブランチ/コミットSHA固定のraw URL（`https://raw.githubusercontent.com/...`）を渡す。投稿前に `curl | md5sum` でキャッシュ齟齬がないか確認すると安全。
- 公開済み投稿のキャプションはAPIでは編集・削除不可（手動のみ）。

## ポスター

- 生成: `python scripts/make_poster.py`（正式版）／`--variant shizuka`（ネタ文言の仮版）。静か版も掲示運用のため `assets/posters/roomradar-poster-shizuka.{png,pdf,pptx}` としてコミット済み。
- `--print-files` でPDF/PPTX（A4縦・約392dpi）も出力。成果物は `assets/posters/` に置く。
- 必須要素: ロゴ・LP QR・Instagram QR（ブランドQR）・「日本大学自主創造プロジェクト」・`#日大生プロジェクト`・TEST表記。
- アプリ画面のスクショを更新する場合: ローカルでapp.pyを起動し、Playwright（390x844, dsf=3）で撮影して `assets/posters/handoff/app-phone-screenshot.png` を差し替え。

## 検証

- 画面系の変更はPlaywright（`/opt/pw-browsers/chromium-1194/chrome-linux/chrome`）でPC幅とスマホ幅（390px）両方のスクショを撮って目視確認する。
- アプリ起動: `PORT=5055 python app.py`（DBは自動生成される）。
