"""
GA4 Data API からサイト分析データを取得し、
data/ga4_history.json（日次の訪問者・PV・セッション推移）と
data/ga4_breakdown.json（イベント数・検索条件・流入元の内訳）を保存する。

dashboard.html がこれらを読み込んで可視化する。

必要な環境変数:
  GA4_PROPERTY_ID : GA4のプロパティID（例 541323442）
  GA4_SA_KEY      : サービスアカウントのJSON鍵（中身そのもの）

サービスアカウントはGA4プロパティに「閲覧者」として追加しておくこと。
"""
import os
import sys
import json
import datetime

try:
    from google.analytics.data_v1beta import BetaAnalyticsDataClient
    from google.analytics.data_v1beta.types import (
        DateRange, Dimension, Metric, RunReportRequest, Filter,
        FilterExpression, FilterExpressionList,
    )
    from google.oauth2 import service_account
except ImportError:
    print("google-analytics-data が未インストールです。pip install google-analytics-data", file=sys.stderr)
    sys.exit(1)

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
HISTORY_FILE = os.path.join(DATA_DIR, "ga4_history.json")
BREAKDOWN_FILE = os.path.join(DATA_DIR, "ga4_breakdown.json")

EVENTS = ["search", "reserve_room", "report_room", "feedback_submit"]


def get_client():
    raw = os.environ.get("GA4_SA_KEY")
    if not raw:
        print("環境変数 GA4_SA_KEY が設定されていません。", file=sys.stderr)
        sys.exit(1)
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/analytics.readonly"])
    return BetaAnalyticsDataClient(credentials=creds)


def run(client, prop, dims, mets, date_range, dim_filter=None, limit=None):
    req = RunReportRequest(
        property=f"properties/{prop}",
        dimensions=[Dimension(name=d) for d in dims],
        metrics=[Metric(name=m) for m in mets],
        date_ranges=[date_range],
        dimension_filter=dim_filter,
        limit=limit,
    )
    return client.run_report(req)


def rows_to_list(resp):
    out = []
    for row in resp.rows:
        out.append({
            "dims": [d.value for d in row.dimension_values],
            "mets": [m.value for m in row.metric_values],
        })
    return out


def main():
    prop = os.environ.get("GA4_PROPERTY_ID")
    if not prop:
        print("環境変数 GA4_PROPERTY_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    client = get_client()
    last28 = DateRange(start_date="28daysAgo", end_date="today")
    os.makedirs(DATA_DIR, exist_ok=True)

    # ── 日次推移（訪問者・新規・PV・セッション） ──
    history = []
    try:
        resp = run(client, prop, ["date"],
                   ["activeUsers", "newUsers", "screenPageViews", "sessions"], last28)
        for r in rows_to_list(resp):
            d = r["dims"][0]
            history.append({
                "date": f"{d[:4]}-{d[4:6]}-{d[6:]}",
                "active_users": int(r["mets"][0]),
                "new_users": int(r["mets"][1]),
                "page_views": int(r["mets"][2]),
                "sessions": int(r["mets"][3]),
            })
        history.sort(key=lambda h: h["date"])
    except Exception as e:
        print(f"[WARN] 日次推移の取得に失敗: {e}", file=sys.stderr)

    breakdown = {"updated": datetime.date.today().isoformat()}

    # ── イベント数 ──
    try:
        ev_filter = FilterExpression(filter=Filter(
            field_name="eventName",
            in_list_filter=Filter.InListFilter(values=EVENTS)))
        resp = run(client, prop, ["eventName"], ["eventCount"], last28, ev_filter)
        breakdown["events"] = {r["dims"][0]: int(r["mets"][0]) for r in rows_to_list(resp)}
    except Exception as e:
        print(f"[WARN] イベント数の取得に失敗: {e}", file=sys.stderr)

    # ── カスタムディメンション別（検索条件） ──
    for key, dim in [("search_building", "customEvent:search_building"),
                     ("search_period", "customEvent:search_period"),
                     ("search_day", "customEvent:search_day")]:
        try:
            resp = run(client, prop, [dim], ["eventCount"], last28, limit=20)
            items = [{"label": r["dims"][0], "count": int(r["mets"][0])}
                     for r in rows_to_list(resp) if r["dims"][0] not in ("(not set)", "")]
            items.sort(key=lambda x: x["count"], reverse=True)
            breakdown[key] = items
        except Exception as e:
            print(f"[WARN] {key} の取得に失敗: {e}", file=sys.stderr)

    # ── 流入元（チャネル） ──
    try:
        resp = run(client, prop, ["sessionDefaultChannelGroup"], ["sessions"], last28, limit=10)
        items = [{"label": r["dims"][0], "sessions": int(r["mets"][0])} for r in rows_to_list(resp)]
        items.sort(key=lambda x: x["sessions"], reverse=True)
        breakdown["channels"] = items
    except Exception as e:
        print(f"[WARN] 流入元の取得に失敗: {e}", file=sys.stderr)

    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    with open(BREAKDOWN_FILE, "w", encoding="utf-8") as f:
        json.dump(breakdown, f, ensure_ascii=False, indent=2)

    print(f"[INFO] {HISTORY_FILE} に {len(history)} 日分を保存")
    print(f"[INFO] {BREAKDOWN_FILE} を保存: {json.dumps(breakdown, ensure_ascii=False)[:200]}")


if __name__ == "__main__":
    main()
