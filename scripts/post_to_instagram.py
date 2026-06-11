"""
Instagram Graph APIへフィード投稿するスクリプト。

必要な環境変数:
  IG_ACCESS_TOKEN : Instagram/Facebookの長期アクセストークン
  IG_ACCOUNT_ID   : InstagramビジネスアカウントID

使い方:
  python post_to_instagram.py --image-url <公開URL> --caption "<キャプション>"
"""
import argparse
import os
import sys
import time
import requests

GRAPH_API_VERSION = "v21.0"
GRAPH_API_BASE = f"https://graph.facebook.com/{GRAPH_API_VERSION}"


def create_media_container(account_id, access_token, image_url, caption):
    url = f"{GRAPH_API_BASE}/{account_id}/media"
    resp = requests.post(url, data={
        "image_url": image_url,
        "caption": caption,
        "access_token": access_token,
    })
    resp.raise_for_status()
    return resp.json()["id"]


def wait_until_ready(container_id, access_token, timeout=120, interval=5):
    url = f"{GRAPH_API_BASE}/{container_id}"
    elapsed = 0
    while elapsed < timeout:
        resp = requests.get(url, params={
            "fields": "status_code",
            "access_token": access_token,
        })
        resp.raise_for_status()
        status = resp.json().get("status_code")
        if status == "FINISHED":
            return
        if status == "ERROR":
            raise RuntimeError("メディアコンテナの処理でエラーが発生しました。")
        time.sleep(interval)
        elapsed += interval
    raise TimeoutError("メディアコンテナの準備がタイムアウトしました。")


def publish_media(account_id, access_token, container_id):
    url = f"{GRAPH_API_BASE}/{account_id}/media_publish"
    resp = requests.post(url, data={
        "creation_id": container_id,
        "access_token": access_token,
    })
    resp.raise_for_status()
    return resp.json()["id"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--image-url", required=True, help="投稿する画像の公開URL")
    parser.add_argument("--caption", required=True, help="投稿のキャプション")
    args = parser.parse_args()

    access_token = os.environ.get("IG_ACCESS_TOKEN")
    account_id = os.environ.get("IG_ACCOUNT_ID")

    if not access_token or not account_id:
        print("環境変数 IG_ACCESS_TOKEN / IG_ACCOUNT_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    print("メディアコンテナを作成しています...")
    container_id = create_media_container(account_id, access_token, args.image_url, args.caption)
    print(f"コンテナID: {container_id}")

    print("コンテナの準備完了を待機しています...")
    wait_until_ready(container_id, access_token)

    print("投稿を公開しています...")
    media_id = publish_media(account_id, access_token, container_id)
    print(f"投稿完了！ メディアID: {media_id}")


if __name__ == "__main__":
    main()
