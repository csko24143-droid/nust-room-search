"""
Instagramの最近の投稿一覧（画像URL・キャプション・投稿日時）を取得して表示するスクリプト。

必要な環境変数:
  IG_ACCESS_TOKEN : Instagram/Facebookの長期アクセストークン
  IG_ACCOUNT_ID   : InstagramビジネスアカウントID
"""
import os
import sys
import requests

GRAPH_API_VERSION = "v21.0"
GRAPH_API_BASE = f"https://graph.facebook.com/{GRAPH_API_VERSION}"


def main():
    access_token = os.environ.get("IG_ACCESS_TOKEN")
    account_id = os.environ.get("IG_ACCOUNT_ID")

    if not access_token or not account_id:
        print("環境変数 IG_ACCESS_TOKEN / IG_ACCOUNT_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    url = f"{GRAPH_API_BASE}/{account_id}/media"
    resp = requests.get(url, params={
        "fields": "id,caption,media_type,media_url,permalink,timestamp",
        "limit": 12,
        "access_token": access_token,
    })
    resp.raise_for_status()
    data = resp.json().get("data", [])

    if not data:
        print("投稿が見つかりませんでした。")
        return

    for item in data:
        print("=" * 60)
        print(f"投稿日時   : {item.get('timestamp')}")
        print(f"メディア種別: {item.get('media_type')}")
        print(f"画像URL    : {item.get('media_url')}")
        print(f"パーマリンク: {item.get('permalink')}")
        print("キャプション:")
        print(item.get("caption", "(なし)"))
        print()


if __name__ == "__main__":
    main()
