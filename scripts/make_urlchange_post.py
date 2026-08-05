#!/usr/bin/env python3
"""URL変更のお知らせ フィード投稿用 1080x1080。
直近投稿（search-update / navbar-update）のハウススタイルに合わせた濃紺デザイン。"""
import base64, pathlib
from playwright.sync_api import sync_playwright

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
ROOT = pathlib.Path(__file__).resolve().parent.parent
STRIP = ROOT / "assets/instagram/handoff/lp-hero-strip.png"
shot = base64.b64encode(pathlib.Path(STRIP).read_bytes()).decode()

HTML = f"""<!doctype html><html><head><meta charset="utf-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700;900&family=Roboto+Mono:wght@500;700&display=swap" rel="stylesheet">
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ width:1080px; height:1080px; overflow:hidden;
  font-family:"Noto Sans JP",sans-serif; color:#fff; }}

/* 濃紺グラデ＋うっすらグリッド（直近投稿と同じ地） */
.bg {{ position:absolute; inset:0;
  background:
    radial-gradient(120% 90% at 82% 8%, rgba(38,76,150,.34) 0%, transparent 60%),
    linear-gradient(150deg,#081326 0%,#0b1a38 48%,#0f2246 100%); }}
.grid {{ position:absolute; inset:0; opacity:.13;
  background-image:
    repeating-linear-gradient(0deg, rgba(150,190,255,.5) 0 1px, transparent 1px 74px),
    repeating-linear-gradient(90deg, rgba(150,190,255,.5) 0 1px, transparent 1px 74px); }}
.glow {{ position:absolute; width:760px; height:760px; right:-230px; top:-250px;
  border-radius:50%; background:radial-gradient(circle,rgba(63,120,235,.22) 0%,transparent 66%); }}

.wrap {{ position:absolute; inset:0; padding:66px 70px 58px;
  display:flex; flex-direction:column; }}

.top {{ display:flex; align-items:center; justify-content:space-between; }}
.badge {{ display:inline-flex; align-items:center; gap:14px;
  padding:13px 30px; border-radius:999px;
  border:1px solid rgba(140,180,255,.42); background:rgba(90,140,235,.14);
  font-family:"Roboto Mono",monospace; font-weight:700; font-size:25px;
  letter-spacing:.18em; color:#a8c6ff; }}
.badge .dot {{ width:13px; height:13px; border-radius:50%; background:#69a1ff;
  box-shadow:0 0 14px rgba(105,161,255,.9); }}
.mark {{ font-size:38px; font-weight:900; letter-spacing:.3px; }}
.mark span {{ color:#6f9ef5; }}

h1 {{ margin-top:40px; font-size:58px; font-weight:900; line-height:1.26; }}
h1 .hl {{ color:#7cadff; }}

.lead {{ margin-top:20px; font-size:24px; font-weight:400; line-height:1.66;
  color:#c2d2ea; }}

/* 白いブラウザ枠（navbar-update / QR投稿の白カードに倣う） */
.browser {{ margin-top:32px; border-radius:20px; overflow:hidden; background:#fff;
  box-shadow:0 26px 60px rgba(0,0,0,.42); }}
.browser .accent {{ height:5px;
  background:linear-gradient(90deg,#2f66f0 0%,#5b8cf5 45%,#f97316 100%); }}
.browser .chrome {{ display:flex; align-items:center; gap:16px;
  padding:13px 18px; background:#f3f5f9; border-bottom:1px solid #e4e8ef; }}
.browser .dots {{ display:flex; gap:8px; }}
.browser .dots i {{ width:13px; height:13px; border-radius:50%; display:block; }}
.browser .addr {{ flex:1; background:#fff; border:1px solid #dfe4ec; border-radius:999px;
  padding:9px 18px; display:flex; align-items:center; gap:11px;
  font-size:21px; font-weight:700; color:#12213f; }}
.browser .addr .lock {{ font-size:19px; }}
.browser .shot {{ display:block; width:100%; }}

.rows {{ margin-top:26px; display:flex; flex-direction:column; gap:11px; }}
.r {{ display:flex; align-items:center; gap:20px;
  background:rgba(255,255,255,.055); border:1px solid rgba(255,255,255,.075);
  border-radius:18px; padding:16px 22px; }}
.r .ico {{ flex:none; width:54px; height:54px; border-radius:14px;
  background:rgba(96,146,240,.2); border:1px solid rgba(130,175,255,.28);
  display:grid; place-items:center; font-size:27px; }}
.r .t {{ font-size:24px; font-weight:700; line-height:1.34; }}
.r .s {{ margin-top:3px; font-size:20px; font-weight:400; color:#a9bcd8; }}

.foot {{ margin-top:auto; display:flex; align-items:flex-end; justify-content:space-between; }}
.foot .disc {{ font-size:20px; line-height:1.5; color:#7f8fa8; }}
.foot .handle {{ font-family:"Roboto Mono",monospace; font-size:27px; font-weight:700;
  color:#7cadff; }}
</style></head>
<body>
<div class="bg"></div>
<div class="grid"></div>
<div class="glow"></div>

<div class="wrap">
  <div class="top">
    <div class="badge"><span class="dot"></span>URL CHANGE</div>
    <div class="mark">Room<span>Radar</span></div>
  </div>

  <h1>URLが<span class="hl">新しく</span><br>なりました</h1>

  <div class="lead">
    運営体制の整備にともない、<br>
    RoomRadarのアドレスを変更しました。
  </div>

  <div class="browser">
    <div class="accent"></div>
    <div class="chrome">
      <div class="dots"><i style="background:#ff5f57"></i><i style="background:#febc2e"></i><i style="background:#28c840"></i></div>
      <div class="addr"><span class="lock">🔒</span>nu-roomradar.github.io/nust-room-search/</div>
    </div>
    <img class="shot" src="data:image/png;base64,{shot}">
  </div>

  <div class="rows">
    <div class="r">
      <div class="ico">🔗</div>
      <div><div class="t">プロフィールのリンクは更新済み</div>
        <div class="s">いつも通りタップするだけでOK</div></div>
    </div>
    <div class="r">
      <div class="ico">⭐</div>
      <div><div class="t">ブックマークしている人は登録し直しを</div>
        <div class="s">古いURLは開けなくなっています</div></div>
    </div>
  </div>

  <div class="foot">
    <div class="disc">RoomRadar ／ テスト運用中の非公式サービスです<br>
      （学生開発・大学公式ではありません）</div>
    <div class="handle">@roomradar_nust</div>
  </div>
</div>
</body></html>"""

out = str(ROOT / "assets/instagram/url-change.png")
with sync_playwright() as p:
    b = p.chromium.launch(executable_path=CHROME)
    pg = b.new_page(viewport={"width": 1080, "height": 1080}, device_scale_factor=2)
    pg.set_content(HTML, wait_until="networkidle")
    pg.wait_for_timeout(600)
    pg.screenshot(path=out, clip={"x": 0, "y": 0, "width": 1080, "height": 1080})
    b.close()
print("saved", out)
