"""
Instagramの最近の投稿一覧（画像URL・キャプション・投稿日時）を取得して表示するスクリプト。

必要な環境変数:
  IG_ACCESS_TOKEN : アクセストークン
  IG_ACCOUNT_ID   : InstagramビジネスアカウントID
"""
import os
import sys
import requests

API_VERSION = "v21.0"
HOSTS = ["https://graph.instagram.com", "https://graph.facebook.com"]
FIELDS = "id,caption,media_type,media_url,permalink,timestamp"


def try_fetch(host, account_id, access_token):
    url = f"{host}/{API_VERSION}/{account_id}/media"
    resp = requests.get(url, params={
        "fields": FIELDS,
        "limit": 12,
        "access_token": access_token,
    })
    return resp


def main():
    access_token = os.environ.get("IG_ACCESS_TOKEN")
    account_id = os.environ.get("IG_ACCOUNT_ID")

    if not access_token or not account_id:
        print("環境変数 IG_ACCESS_TOKEN / IG_ACCOUNT_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    data = None
    for host in HOSTS:
        resp = try_fetch(host, account_id, access_token)
        if resp.ok:
            print(f"[INFO] 使用エンドポイント: {host}")
            data = resp.json().get("data", [])
            break
        # トークンを含まない形でエラー内容を表示
        try:
            err = resp.json().get("error", {})
            print(f"[WARN] {host} -> {resp.status_code}: "
                  f"type={err.get('type')} code={err.get('code')} message={err.get('message')}",
                  file=sys.stderr)
        except Exception:
            print(f"[WARN] {host} -> {resp.status_code}（本文の解析に失敗）", file=sys.stderr)

    if data is None:
        # アカウントIDが違う可能性に備えて /me も試す
        for host in HOSTS:
            resp = requests.get(f"{host}/{API_VERSION}/me", params={
                "fields": "id,username",
                "access_token": access_token,
            })
            if resp.ok:
                me = resp.json()
                print(f"[INFO] /me ({host}) -> id={me.get('id')} username={me.get('username')}")
                print("[INFO] IG_ACCOUNT_ID をこのIDに更新すると解決する可能性があります。")
            else:
                try:
                    err = resp.json().get("error", {})
                    print(f"[WARN] /me {host} -> {resp.status_code}: {err.get('message')}", file=sys.stderr)
                except Exception:
                    pass
        sys.exit(1)

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
