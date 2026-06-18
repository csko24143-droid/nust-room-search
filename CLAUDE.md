# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

RoomRadar (日大理工学部 空き教室検索) — a real-time empty-classroom finder for Nihon University's
College of Science and Technology. Students pick a day/period/building and see which classrooms
have no scheduled class, plus a lightweight "soft reservation" / "actually in use" reporting layer
on top.

The repo contains three independently-deployed surfaces that share branding but not code:

1. **`app.py`** — the actual Flask application (search + reservation/report API), deployed to
   Render at `https://nust-room-search.onrender.com`.
2. **`index.html`** — a static marketing/landing page deployed via GitHub Pages
   (`csko24143-droid.github.io/nust-room-search`). It links out to the Render-hosted app for the
   real search feature and is multi-language (ja/en/zh, toggled via CSS classes — see below).
3. **`dashboard.html`** — a static, `noindex` analytics dashboard (Chart.js) that reads the JSON
   files in `data/` to show Instagram/GA insights. Not linked from the main UI.

There is no shared build system between these — each is a self-contained HTML file (or, for
`app.py`, a Python file with the HTML/CSS/JS embedded as a Python string template).

## Commands

No test suite, linter, or build step exists in this repo.

```bash
pip install -r requirements.txt   # flask, pandas, openpyxl, gunicorn

python app.py                     # run the Flask app locally (reads PORT env var, default 10000)
gunicorn app:app                  # production entrypoint (used by Render)
```

Instagram automation scripts (run manually or via the GitHub Actions workflows in
`.github/workflows/`) require `IG_ACCESS_TOKEN` and `IG_ACCOUNT_ID` env vars:

```bash
python scripts/fetch_instagram_posts.py
python scripts/fetch_instagram_insights.py
python scripts/post_to_instagram.py --image-url <url> --caption "<text>"
```

## Architecture: `app.py`

Single-file Flask app. Everything — routes, DB access, and the page itself
(`HTML_TEMPLATE`, a Jinja string rendered via `render_template_string`) — lives in this one file.
There is no `templates/` or `static/` directory.

**Three separate SQLite databases, with different lifecycles:**

- `schedule_final.db` — checked into git, treated as read-only source-of-truth data. Has two
  tables: `schedules` (course timetable, Japanese column names: 学科/履修期名/曜日/時限/教室/校舎/科目名)
  and `classrooms` (`name`, `building`). `classroom_data.xlsx` / `summry_classrooms.xlsx` are the
  raw spreadsheets this DB was built from — they are not read by any code at runtime, so if the
  timetable needs updating, `schedule_final.db` itself must be regenerated/replaced directly.
- `reservations.db` / `reports.db` — created at runtime by `init_reserve_db()` /
  `init_reports_db()`, gitignored. These back the "soft reservation" and "actually in use" report
  features. Both use a `cancel_code` pattern: an anonymous random 6-char code is returned to the
  client on create and stored in `localStorage`; the same code is required to delete/cancel later.
  There is no auth — anyone with the code can cancel.

**Term/time logic:** `ACTIVE_TERMS` and `get_active_terms()` decide whether 前期 (spring) or 後期
(fall) schedule rows apply, based on a hardcoded month/day window (4/1–9/20 = 前期). `PERIODS`
maps period number (1–6) to start/end clock times in JST; `period_end_dt()` uses this to compute
when a reservation/report should auto-expire (cleaned up by `cleanup_expired()` /
`cleanup_reports()`, called at the top of relevant requests — there is no background job).

**Rate limiting and validation are in-process and in-memory** (`_rate_store`, a
`collections.defaultdict(list)` keyed by IP). This only works correctly with a single worker
process — if Render/gunicorn is ever scaled to multiple workers, rate limiting will be
per-worker, not global. `VALID_DAYS`/`VALID_PERIODS`/`VALID_BUILDINGS` are the whitelist for all
incoming request params.

**API routes** (`/api/reserve`, `/api/reserve/cancel`, `/api/reserve/list`, `/api/report`,
`/api/report/cancel`) are plain JSON POST/GET endpoints consumed by inline `<script>` in
`HTML_TEMPLATE`; `/` (GET/POST) is the search page itself, POST being a search submission.

## Architecture: Instagram automation (`scripts/` + `.github/workflows/`)

Three scripts wrap the Instagram Graph API, each trying `graph.instagram.com` then
`graph.facebook.com` as fallback hosts (`HOSTS` list pattern repeated in all three files):

- `fetch_instagram_posts.py` — debug/inspection script, prints recent posts to stdout. Manual
  `workflow_dispatch` only.
- `fetch_instagram_insights.py` — the real data pipeline: appends a daily snapshot to
  `data/instagram_history.json` and overwrites `data/instagram_posts.json` with latest post
  performance. Runs on a cron (Mon/Thu 09:00 JST) via `instagram-insights.yml`, which then commits
  the updated JSON straight back to the branch as `github-actions[bot]`. `dashboard.html` reads
  these two JSON files client-side to render charts.
- `post_to_instagram.py` — publishes a new feed post given an image URL + caption. Manual
  `workflow_dispatch` with `image_url`/`caption` inputs only; never run automatically.

If you change the shape of `data/instagram_history.json` or `data/instagram_posts.json`, update
both the writer (`fetch_instagram_insights.py`) and the reader (`dashboard.html`'s JS) together.

## Conventions

- Database/table/column identifiers, and most server error/log strings, are Japanese — keep new
  fields consistent with the existing naming rather than introducing English equivalents.
- All datetime handling goes through `JST` (`datetime.timezone(timedelta(hours=9))`); don't use
  naive `datetime.now()` for anything user-facing.
- `index.html`'s language toggle works by adding a `lang-en`/`lang-zh` class to `<body>` and
  relying on CSS (`.ja`/`.en`/`.zh` block-level, `span.ja`/`span.en`/`span.zh` inline) to show/hide
  pre-rendered translated copy — there's no i18n library or runtime string lookup. New copy needs
  all three language variants added inline, following the existing pattern.
