#!/usr/bin/env python3
"""紹介お礼ストーリー（@nu_su_project に紹介された件のリポスト用）1080x1920。"""
import base64, pathlib
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

def b64(p):
    return base64.b64encode(pathlib.Path(p).read_bytes()).decode()

logo = b64(ROOT / "rr_logo.png")
shot = b64(ROOT / "assets/instagram/handoff/nusu-intro-screenshot.png")

HTML = f"""<!doctype html><html><head><meta charset="utf-8"><style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
html,body {{ width:1080px; height:1920px; }}
body {{
  font-family:'Hiragino Kaku Gothic ProN','Noto Sans JP','Noto Sans CJK JP',system-ui,sans-serif;
  position:relative; overflow:hidden;
  background:
    radial-gradient(760px 620px at 14% 8%, rgba(255,197,140,.55), transparent 60%),
    radial-gradient(720px 640px at 92% 96%, rgba(120,160,255,.50), transparent 60%),
    linear-gradient(158deg,#fff4ea 0%,#fdeef4 44%,#e9f0ff 78%,#dde8ff 100%);
  color:#12213f;
}}
/* 上端のIGユーザー名バー/下端の返信バーに隠れないよう安全マージンを確保 */
.wrap {{ position:absolute; inset:0; padding:92px 76px 190px; display:flex; flex-direction:column; }}

/* header */
.brand {{ display:flex; align-items:center; gap:18px; }}
.brand img {{ width:82px; height:82px; object-fit:contain;
  filter:drop-shadow(0 6px 14px rgba(20,40,80,.18)); }}
.brand .wm {{ font-size:42px; font-weight:800; letter-spacing:.3px; }}
.brand .wm .r {{ color:#2f66f0; }}
.brand .tag {{ font-size:21px; font-weight:600; color:#5b6b86; margin-top:3px; }}

/* headline block */
.eyebrow {{ display:inline-flex; align-items:center; gap:12px; align-self:flex-start;
  margin-top:30px; padding:14px 28px; border-radius:999px;
  background:#12213f; color:#fff; font-size:29px; font-weight:700;
  box-shadow:0 10px 26px rgba(18,33,63,.28); }}
h1 {{ margin-top:16px; font-size:74px; font-weight:900; line-height:1.13; letter-spacing:.5px; }}
h1 .hl {{ color:#2f66f0; }}
.sub {{ margin-top:16px; font-size:32px; font-weight:700; line-height:1.5; color:#28374f; }}
.sub .mention {{ color:#2f66f0; font-weight:800; }}

/* screenshot card */
.cardwrap {{ flex:1; min-height:0; display:flex; align-items:flex-start; justify-content:center; margin:58px 0 0; }}
.card {{ position:relative; transform:rotate(-2.4deg);
  background:#fff; padding:16px 16px 20px; border-radius:26px;
  box-shadow:0 34px 70px rgba(20,40,90,.34), 0 8px 20px rgba(20,40,90,.20);
  border:1px solid rgba(255,255,255,.9); }}
.card img {{ display:block; width:532px; border-radius:14px; }}
.stamp {{ position:absolute; top:-30px; right:-30px; transform:rotate(6deg);
  background:#ff5a7a; color:#fff; font-size:33px; font-weight:800;
  padding:15px 28px; border-radius:16px;
  box-shadow:0 14px 30px rgba(255,90,122,.44); }}

/* footer */
.thanks {{ font-size:37px; font-weight:800; text-align:center; letter-spacing:.5px; }}
.thanks .em {{ color:#2f66f0; }}
.handle {{ text-align:center; margin-top:18px; font-size:31px; font-weight:800;
  color:#12213f; letter-spacing:.5px; }}
</style></head><body>
<div class="wrap">
  <div class="brand">
    <img src="data:image/png;base64,{logo}">
    <div>
      <div class="wm">Room<span class="r">Radar</span></div>
      <div class="tag">日本大学理工学部・空き教室リアルタイム検索</div>
    </div>
  </div>

  <div class="eyebrow">🎉 うれしいご報告</div>
  <h1>ご紹介、<br><span class="hl">いただきました！</span></h1>
  <div class="sub"><span class="mention">@nu_su_project</span> さんが RoomRadar を<br>理工学部のみなさんに紹介してくれました🙏</div>

  <div class="cardwrap">
    <div class="card">
      <div class="stamp">Thank you!</div>
      <img src="data:image/png;base64,{shot}">
    </div>
  </div>

  <div class="thanks">あたたかい<span class="em">ご紹介、ありがとうございます</span></div>
  <div class="handle">@roomradar_nust</div>
</div>
</body></html>"""

out = str(ROOT / "assets/instagram/intro-thanks-story.png")
with sync_playwright() as p:
    b = p.chromium.launch(executable_path=CHROME)
    pg = b.new_page(viewport={"width":1080,"height":1920}, device_scale_factor=2)
    pg.set_content(HTML, wait_until="networkidle")
    pg.wait_for_timeout(400)
    pg.screenshot(path=out, clip={"x":0,"y":0,"width":1080,"height":1920})
    b.close()
print("saved", out)
