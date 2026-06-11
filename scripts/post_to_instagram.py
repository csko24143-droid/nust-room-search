"""
Instagram Graph APIへフィード投稿するスクリプト。

必要な環境変数:
  IG_ACCESS_TOKEN : アクセストークン
  IG_ACCOUNT_ID   : InstagramビジネスアカウントID

使い方:
  python post_to_instagram.py --image-url <公開URL> --caption "<キャプション>"
"""
import argparse
import os
import sys
import time
import requests

API_VERSION = "v21.0"
HOSTS = ["https://graph.instagram.com", "https://graph.facebook.com"]


def show_error(resp, context):
    try:
        err = resp.json().get("error", {})
        print(f"[ERROR] {context} -> {resp.status_code}: "
              f"type={err.get('type')} code={err.get('code')} message={err.get('message')}",
              file=sys.stderr)
    except Exception:
        print(f"[ERROR] {context} -> {resp.status_code}（本文の解析に失敗）", file=sys.stderr)


def detect_host(account_id, access_token):
    """利用可能なAPIホストを判定する"""
    for host in HOSTS:
        resp = requests.get(f"{host}/{API_VERSION}/{account_id}", params={
            "fields": "id",
            "access_token": access_token,
        })
        if resp.ok:
            print(f"[INFO] 使用エンドポイント: {host}")
            return host
        show_error(resp, f"ホスト判定 {host}")
    raise RuntimeError("どのAPIホストでもアカウントにアクセスできませんでした。")


def create_media_container(host, account_id, access_token, image_url, caption):
    resp = requests.post(f"{host}/{API_VERSION}/{account_id}/media", data={
        "image_url": image_url,
        "caption": caption,
        "access_token": access_token,
    })
    if not resp.ok:
        show_error(resp, "コンテナ作成")
        resp.raise_for_status()
    return resp.json()["id"]


def wait_until_ready(host, container_id, access_token, timeout=120, interval=5):
    elapsed = 0
    while elapsed < timeout:
        resp = requests.get(f"{host}/{API_VERSION}/{container_id}", params={
            "fields": "status_code",
            "access_token": access_token,
        })
        if not resp.ok:
            show_error(resp, "ステータス確認")
            resp.raise_for_status()
        status = resp.json().get("status_code")
        if status == "FINISHED":
            return
        if status == "ERROR":
            raise RuntimeError("メディアコンテナの処理でエラーが発生しました。")
        time.sleep(interval)
        elapsed += interval
    raise TimeoutError("メディアコンテナの準備がタイムアウトしました。")


def publish_media(host, account_id, access_token, container_id):
    resp = requests.post(f"{host}/{API_VERSION}/{account_id}/media_publish", data={
        "creation_id": container_id,
        "access_token": access_token,
    })
    if not resp.ok:
        show_error(resp, "公開")
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

    host = detect_host(account_id, access_token)

    print("メディアコンテナを作成しています...")
    container_id = create_media_container(host, account_id, access_token, args.image_url, args.caption)
    print(f"コンテナID: {container_id}")

    print("コンテナの準備完了を待機しています...")
    wait_until_ready(host, container_id, access_token)

    print("投稿を公開しています...")
    media_id = publish_media(host, account_id, access_token, container_id)
    print(f"投稿完了！ メディアID: {media_id}")


if __name__ == "__main__":
    main()
