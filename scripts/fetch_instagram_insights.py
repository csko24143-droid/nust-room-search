"""
Instagramアカウント・投稿のインサイト（分析データ）を取得し、
data/instagram_history.json に日次スナップショットを追記、
data/instagram_posts.json に最新投稿のパフォーマンスを保存する。

ダッシュボード(dashboard.html)がこれらのJSONを読み込んで可視化する。

必要な環境変数:
  IG_ACCESS_TOKEN : アクセストークン
  IG_ACCOUNT_ID   : InstagramビジネスアカウントID
"""
import os
import sys
import json
import datetime
import requests

API_VERSION = "v21.0"
HOSTS = ["https://graph.instagram.com", "https://graph.facebook.com"]

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
HISTORY_FILE = os.path.join(DATA_DIR, "instagram_history.json")
POSTS_FILE = os.path.join(DATA_DIR, "instagram_posts.json")

# アカウント基本フィールド（常に取得可能）
ACCOUNT_FIELDS = "username,name,followers_count,follows_count,media_count"
# アカウント単位のインサイト（APIバージョンで可否が変わるため個別にトライ）
ACCOUNT_METRICS = ["reach", "profile_views", "accounts_engaged", "total_interactions"]
# 投稿単位のインサイト
MEDIA_METRICS = "likes,comments,saved,shares,reach,views,total_interactions"
MEDIA_FIELDS = "id,caption,media_type,media_url,thumbnail_url,permalink,timestamp,like_count,comments_count"


def pick_host(account_id, access_token):
    """有効なホストを判定して返す。"""
    for host in HOSTS:
        resp = requests.get(f"{host}/{API_VERSION}/{account_id}",
                            params={"fields": "id", "access_token": access_token})
        if resp.ok:
            print(f"[INFO] 使用エンドポイント: {host}")
            return host
        try:
            err = resp.json().get("error", {})
            print(f"[WARN] {host} -> {resp.status_code}: {err.get('message')}", file=sys.stderr)
        except Exception:
            print(f"[WARN] {host} -> {resp.status_code}", file=sys.stderr)
    return None


def get_account(host, account_id, token):
    resp = requests.get(f"{host}/{API_VERSION}/{account_id}",
                        params={"fields": ACCOUNT_FIELDS, "access_token": token})
    resp.raise_for_status()
    return resp.json()


def get_account_insights(host, account_id, token):
    """アカウント単位のインサイトを個別に取得（失敗したメトリクスはスキップ）。"""
    results = {}
    for metric in ACCOUNT_METRICS:
        params = {"metric": metric, "period": "day",
                  "metric_type": "total_value", "access_token": token}
        resp = requests.get(f"{host}/{API_VERSION}/{account_id}/insights", params=params)
        if not resp.ok:
            # metric_type 無しでも再トライ
            params.pop("metric_type", None)
            resp = requests.get(f"{host}/{API_VERSION}/{account_id}/insights", params=params)
        if not resp.ok:
            try:
                print(f"[WARN] insight {metric} -> {resp.json().get('error', {}).get('message')}",
                      file=sys.stderr)
            except Exception:
                pass
            continue
        try:
            entry = resp.json().get("data", [])[0]
            tv = entry.get("total_value")
            if tv is not None:
                results[metric] = tv.get("value")
            else:
                values = entry.get("values", [])
                if values:
                    results[metric] = values[-1].get("value")
        except (IndexError, KeyError, TypeError):
            continue
    return results


def get_media(host, account_id, token, limit=12):
    resp = requests.get(f"{host}/{API_VERSION}/{account_id}/media",
                        params={"fields": MEDIA_FIELDS, "limit": limit, "access_token": token})
    resp.raise_for_status()
    return resp.json().get("data", [])


def get_media_insights(host, media_id, token):
    resp = requests.get(f"{host}/{API_VERSION}/{media_id}/insights",
                        params={"metric": MEDIA_METRICS, "access_token": token})
    if not resp.ok:
        return {}
    out = {}
    for item in resp.json().get("data", []):
        try:
            out[item["name"]] = item["values"][0]["value"]
        except (KeyError, IndexError):
            pass
    return out


def main():
    token = os.environ.get("IG_ACCESS_TOKEN")
    account_id = os.environ.get("IG_ACCOUNT_ID")
    if not token or not account_id:
        print("環境変数 IG_ACCESS_TOKEN / IG_ACCOUNT_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    host = pick_host(account_id, token)
    if not host:
        print("有効なエンドポイントが見つかりませんでした。", file=sys.stderr)
        sys.exit(1)

    today = datetime.date.today().isoformat()
    account = get_account(host, account_id, token)
    insights = get_account_insights(host, account_id, token)

    snapshot = {
        "date": today,
        "followers_count": account.get("followers_count"),
        "follows_count": account.get("follows_count"),
        "media_count": account.get("media_count"),
        **insights,
    }
    print("[INFO] スナップショット:", json.dumps(snapshot, ensure_ascii=False))

    # 投稿パフォーマンス
    posts = []
    for m in get_media(host, account_id, token):
        mi = get_media_insights(host, m["id"], token)
        posts.append({
            "id": m.get("id"),
            "permalink": m.get("permalink"),
            "media_type": m.get("media_type"),
            "thumbnail": m.get("thumbnail_url") or m.get("media_url"),
            "caption": m.get("caption") or "",
            "timestamp": m.get("timestamp"),
            "like_count": m.get("like_count"),
            "comments_count": m.get("comments_count"),
            "saved": mi.get("saved"),
            "shares": mi.get("shares"),
            "reach": mi.get("reach"),
            "views": mi.get("views"),
            "total_interactions": mi.get("total_interactions"),
        })

    os.makedirs(DATA_DIR, exist_ok=True)

    # 履歴に追記（同じ日付があれば上書き）
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, encoding="utf-8") as f:
                history = json.load(f)
        except (json.JSONDecodeError, OSError):
            history = []
    history = [h for h in history if h.get("date") != today]
    history.append(snapshot)
    history.sort(key=lambda h: h.get("date", ""))

    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    with open(POSTS_FILE, "w", encoding="utf-8") as f:
        json.dump({"updated": today, "posts": posts}, f, ensure_ascii=False, indent=2)

    print(f"[INFO] {HISTORY_FILE} に {len(history)} 件の履歴を保存")
    print(f"[INFO] {POSTS_FILE} に {len(posts)} 件の投稿を保存")


if __name__ == "__main__":
    main()
