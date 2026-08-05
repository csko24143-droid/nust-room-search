"""
フィード投稿の画像から「新着投稿のお知らせ」ストーリーズ画像(1080x1920)を生成する。

投稿した正方形画像をカードとして埋め込み、「新しい投稿をしました」＋
「プロフィールからチェック」の導線を重ねたストーリーズを作る。

使い方:
  python make_story_promo.py --feed-image-url <公開URL> --out <出力PNGパス>
  # 見出しは --headline で明示するか、--caption を渡すと投稿内容から自動生成される
  python make_story_promo.py --feed-image-url <URL> --out <PNG> --caption "$CAPTION"

見出しの自動生成にはClaude APIを使う（環境変数 ANTHROPIC_API_KEY）。
キーが無い場合やAPIが失敗した場合は、キャプションの1行目にフォールバックするため、
投稿自体は止まらない。

Instagramの仕様上、API経由のストーリーズには投稿への
タップ可能なリンクスタンプを付けられないため、画像内の文言で誘導する。
"""
import argparse
import base64
import html as html_mod
import os
import sys
import requests
from playwright.sync_api import sync_playwright

DEFAULT_HEADLINE = "新しい投稿をしました📣"

TEMPLATE = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Noto+Sans+JP:wght@400;500;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  *{margin:0;padding:0;box-sizing:border-box;}
  body{width:1080px;height:1920px;font-family:'Noto Sans JP',sans-serif;position:relative;overflow:hidden;
    background:linear-gradient(160deg,#0c1524 0%,#13233a 55%,#0c1524 100%);}
  .grid{position:absolute;inset:0;opacity:0.5;
    background-image:linear-gradient(rgba(255,255,255,0.05) 1px,transparent 1px),
      linear-gradient(90deg,rgba(255,255,255,0.05) 1px,transparent 1px);background-size:54px 54px;}
  .blob{position:absolute;border-radius:50%;filter:blur(120px);opacity:0.30;}
  .b1{width:560px;height:560px;background:#2563eb;top:-180px;right:-160px;}
  .b2{width:440px;height:440px;background:#f97316;bottom:-160px;left:-140px;opacity:0.20;}
  .wrap{position:relative;z-index:1;padding:150px 96px;height:100%;display:flex;flex-direction:column;align-items:center;}
  .brand{font-family:'Syne',sans-serif;font-weight:800;font-size:42px;color:#fff;}
  .brand b{color:#7dd3fc;}
  .eyebrow{display:inline-flex;align-items:center;gap:12px;margin-top:84px;
    font-family:'DM Mono',monospace;font-size:28px;letter-spacing:0.12em;text-transform:uppercase;
    color:#7dd3fc;background:rgba(125,211,252,0.1);border:1px solid rgba(125,211,252,0.3);
    padding:16px 32px;border-radius:99px;}
  .eyebrow .dot{width:13px;height:13px;border-radius:50%;background:#7dd3fc;}
  h1{font-weight:800;font-size:88px;line-height:1.3;color:#fff;margin-top:48px;text-align:center;
    width:100%;overflow-wrap:break-word;text-wrap:balance;}
  h1 .hl{color:#7dd3fc;}
  .photo{margin-top:64px;width:680px;height:680px;border-radius:40px;overflow:hidden;
    box-shadow:0 30px 70px rgba(0,0,0,0.5);border:6px solid rgba(255,255,255,0.9);position:relative;}
  .photo img{width:100%;height:100%;object-fit:cover;display:block;}
  .tag{position:absolute;top:28px;left:28px;font-family:'DM Mono',monospace;font-size:26px;font-weight:700;
    color:#fff;background:#2563eb;padding:12px 24px;border-radius:99px;box-shadow:0 8px 20px rgba(37,99,235,0.4);}
  .cta{margin-top:auto;display:flex;flex-direction:column;align-items:center;gap:26px;}
  .swipe{font-family:'DM Mono',monospace;font-size:32px;font-weight:700;letter-spacing:0.06em;color:#fff;
    background:#2563eb;padding:28px 60px;border-radius:99px;box-shadow:0 16px 40px rgba(37,99,235,0.4);
    display:inline-flex;align-items:center;gap:16px;}
  .swipe .ar{font-size:36px;}
  .handle{font-family:'DM Mono',monospace;font-size:30px;color:#7dd3fc;letter-spacing:0.04em;}
</style></head>
<body>
  <div class="grid"></div><div class="blob b1"></div><div class="blob b2"></div>
  <div class="wrap">
    <div class="brand">Room<b>Radar</b></div>
    <div class="eyebrow"><span class="dot"></span>NEW POST</div>
    <h1>__HEADLINE__</h1>
    <div class="photo"><span class="tag">🆕 NEW</span><img src="__IMG__"></div>
    <div class="cta">
      <span class="swipe">プロフィールでチェック<span class="ar">→</span></span>
      <span class="handle">@roomradar_nust</span>
    </div>
  </div>
</body></html>"""


HEADLINE_SYSTEM = """あなたはInstagramの運用担当者です。
フィード投稿のキャプションを読み、それを告知するストーリーの見出しを1つ作ってください。

条件:
- 投稿の内容が一目で伝わる、具体的な見出しにする
- 全角14文字以内。長いと画像からはみ出すため厳守する
- 文頭に内容に合った絵文字を1つ付ける
- 「新しい投稿をしました」のような、内容を説明しない汎用的な文言にはしない
- キャプションの1行目をそのまま写すのではなく、ストーリーの見出しとして最適な長さと言い回しに整える
- 体言止めまたは「〜しました」程度の簡潔な形にする。説明文にはしない
- 弔事・お見舞いなど不幸に関する内容の場合は、絵文字を付けず、簡素な表現にする

見出しだけを出力し、前置きや説明、鉤括弧は付けないこと。"""


def headline_from_caption(caption):
    """キャプションの1行目。API未使用時のフォールバック。"""
    for line in (caption or "").splitlines():
        line = line.strip()
        if line:
            return line
    return DEFAULT_HEADLINE


def generate_headline(caption):
    """投稿内容をClaudeに読ませて見出しを生成する。

    失敗しても投稿を止めないよう、例外は握りつぶして1行目にフォールバックする。
    """
    fallback = headline_from_caption(caption)
    if not (caption or "").strip():
        return fallback
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("[INFO] ANTHROPIC_API_KEY が未設定のため、キャプションの1行目を見出しに使います。")
        return fallback

    try:
        import anthropic
    except ImportError:
        print("[WARN] anthropic パッケージが無いため、キャプションの1行目を見出しに使います。",
              file=sys.stderr)
        return fallback

    try:
        client = anthropic.Anthropic()
        response = client.messages.create(
            model="claude-opus-5",
            max_tokens=200,
            system=HEADLINE_SYSTEM,
            output_config={"effort": "low"},
            messages=[{"role": "user", "content": f"<caption>\n{caption}\n</caption>"}],
        )
        if response.stop_reason == "refusal":
            print("[WARN] 見出し生成が拒否されました。キャプションの1行目を使います。", file=sys.stderr)
            return fallback
        text = "".join(b.text for b in response.content if b.type == "text").strip()
        headline = text.splitlines()[0].strip().strip("「」\"'") if text else ""
        if not headline:
            print("[WARN] 見出しが空でした。キャプションの1行目を使います。", file=sys.stderr)
            return fallback
        return headline
    except Exception as exc:
        print(f"[WARN] 見出しの自動生成に失敗（{type(exc).__name__}）。"
              f"キャプションの1行目を使います。", file=sys.stderr)
        return fallback


# 見出しの級数決定と最終行のハイライトは、ブラウザ上で実測して行う。
# 文字数から機械的に見積もると、太字の欧文や絵文字の幅がずれて崩れるため。
FIT_HEADLINE_JS = """
(text) => {
  const h1 = document.querySelector('h1');
  const sizes = [88, 78, 68, 60, 54];
  let chosen = sizes[sizes.length - 1];
  // まず2行に収まる級数を探す（3行になるとカードとCTAが窮屈になるため）
  for (const s of sizes) {
    h1.style.fontSize = s + 'px';
    h1.textContent = text;
    if (h1.getBoundingClientRect().height <= s * 1.3 * 2 + 2) { chosen = s; break; }
  }
  // どの級数でも2行に収まらない長文だけ、最小級数で3行を許容する
  h1.style.fontSize = chosen + 'px';
  h1.textContent = text;
  if (h1.getBoundingClientRect().height > chosen * 1.3 * 2 + 2) {
    chosen = sizes[sizes.length - 1];
  }
  h1.style.fontSize = chosen + 'px';
  h1.textContent = text;

  // 各文字の描画位置を見て、最終行が始まるインデックスを割り出す
  const node = h1.firstChild;
  const range = document.createRange();
  let maxTop = -Infinity, startIdx = 0;
  for (let i = 0; i < text.length; i++) {
    range.setStart(node, i); range.setEnd(node, i + 1);
    const top = Math.round(range.getBoundingClientRect().top);
    if (top > maxTop) { maxTop = top; startIdx = i; }
  }

  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const head = text.slice(0, startIdx), tail = text.slice(startIdx);
  h1.innerHTML = esc(head) + '<span class="hl">' + esc(tail) + '</span>';
  return { size: chosen, lines: [head, tail] };
}
"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--feed-image-url", required=True, help="埋め込むフィード画像の公開URL")
    parser.add_argument("--out", required=True, help="出力PNGのパス")
    parser.add_argument("--headline", default=None,
                        help="ストーリーの見出し。指定するとClaudeによる自動生成をスキップする")
    parser.add_argument("--caption", default=None,
                        help="フィード投稿のキャプション。内容を読んで見出しを自動生成する")
    args = parser.parse_args()

    headline = (args.headline or "").strip() or generate_headline(args.caption)
    headline = headline.strip() or DEFAULT_HEADLINE

    resp = requests.get(args.feed_image_url, timeout=30)
    if not resp.ok:
        print(f"[ERROR] フィード画像の取得に失敗: {resp.status_code}", file=sys.stderr)
        sys.exit(1)
    b64 = base64.b64encode(resp.content).decode()
    data_uri = f"data:image/png;base64,{b64}"
    html = TEMPLATE.replace("__IMG__", data_uri).replace("__HEADLINE__", html_mod.escape(headline))
    print(f"[INFO] 見出し: {headline}")

    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox"])
        page = browser.new_page(viewport={"width": 1080, "height": 1920})
        page.set_content(html, wait_until="networkidle")
        page.wait_for_timeout(800)
        fitted = page.evaluate(FIT_HEADLINE_JS, headline)
        print(f"[INFO] 見出しの級数: {fitted['size']}px")
        page.wait_for_timeout(200)
        page.screenshot(path=args.out)
        browser.close()
    print(f"[INFO] ストーリーズ画像を生成: {args.out}")


if __name__ == "__main__":
    main()
