"""
Instagram Graph APIへフィード投稿するスクリプト。

必要な環境変数:
  IG_ACCESS_TOKEN : アクセストークン
  IG_ACCOUNT_ID   : InstagramビジネスアカウントID

使い方:
  # フィード投稿
  python post_to_instagram.py --image-url <公開URL> --caption "<キャプション>"
  # ストーリーズ投稿（キャプションは任意・ストーリーズには表示されない）
  python post_to_instagram.py --image-url <公開URL> --story
"""
import argparse
import os
import sys
import time
import requests

API_VERSION = "v21.0"
HOSTS = ["https://graph.instagram.com", "https://graph.facebook.com"]
# 自主創造プロジェクトのレギュレーションで学外SNS発信時の使用が義務付けられた共通ハッシュタグ
REQUIRED_TAGS = ["#日大生プロジェクト"]


def ensure_required_tags(caption):
    missing = [tag for tag in REQUIRED_TAGS if tag not in caption]
    if not missing:
        return caption
    return caption.rstrip() + "\n" + " ".join(missing)


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


def create_media_container(host, account_id, access_token, image_url, caption, story=False):
    data = {
        "image_url": image_url,
        "access_token": access_token,
    }
    if story:
        # ストーリーズはキャプション欄を持たないため caption は付けない
        data["media_type"] = "STORIES"
    else:
        data["caption"] = caption
    resp = requests.post(f"{host}/{API_VERSION}/{account_id}/media", data=data)
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
    parser.add_argument("--caption", default="", help="投稿のキャプション（フィードのみ）")
    parser.add_argument("--story", action="store_true", help="ストーリーズとして投稿する")
    args = parser.parse_args()

    if not args.story and not args.caption:
        print("フィード投稿には --caption が必要です。", file=sys.stderr)
        sys.exit(1)

    access_token = os.environ.get("IG_ACCESS_TOKEN")
    account_id = os.environ.get("IG_ACCOUNT_ID")

    if not access_token or not account_id:
        print("環境変数 IG_ACCESS_TOKEN / IG_ACCOUNT_ID が設定されていません。", file=sys.stderr)
        sys.exit(1)

    host = detect_host(account_id, access_token)

    caption = args.caption if args.story else ensure_required_tags(args.caption)

    kind = "ストーリーズ" if args.story else "フィード"
    print(f"メディアコンテナを作成しています...（{kind}）")
    container_id = create_media_container(
        host, account_id, access_token, args.image_url, caption, story=args.story)
    print(f"コンテナID: {container_id}")

    print("コンテナの準備完了を待機しています...")
    wait_until_ready(host, container_id, access_token)

    print("投稿を公開しています...")
    media_id = publish_media(host, account_id, access_token, container_id)
    print(f"投稿完了！ メディアID: {media_id}")


if __name__ == "__main__":
    main()
