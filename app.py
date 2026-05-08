import os
import sqlite3
import datetime
from flask import Flask, render_template_string, request

app = Flask(__name__)

DB_NAME = "schedule_final.db"
JST = datetime.timezone(datetime.timedelta(hours=9))
PERIODS = {1: ("09:00", "10:40"), 2: ("10:50", "12:30"), 3: ("13:20", "15:00"),
           4: ("15:10", "16:50"), 5: ("17:00", "18:40"), 6: ("18:50", "20:30")}
ACTIVE_TERMS = ['前期', '通年']

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
    <title>RoomRadar — 日大理工 空き教室</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg: #0a0a0f;
            --surface: #13131a;
            --surface2: #1c1c26;
            --border: rgba(255,255,255,0.07);
            --text: #f0f0f5;
            --muted: #6b6b80;
            --tower-color: #6c8fff;
            --tower-bg: rgba(108,143,255,0.08);
            --tower-border: rgba(108,143,255,0.3);
            --surugadai-color: #4dd9a0;
            --surugadai-bg: rgba(77,217,160,0.08);
            --surugadai-border: rgba(77,217,160,0.3);
            --funabashi-color: #ff8c5a;
            --funabashi-bg: rgba(255,140,90,0.08);
            --funabashi-border: rgba(255,140,90,0.3);
            --accent: #6c8fff;
            --radius: 14px;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Noto Sans JP', sans-serif;
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
            padding: 0 0 env(safe-area-inset-bottom) 0;
        }

        /* ── ノイズ背景 ── */
        body::before {
            content: '';
            position: fixed; inset: 0;
            background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='noise'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23noise)' opacity='0.03'/%3E%3C/svg%3E");
            pointer-events: none; z-index: 0;
        }

        .wrap { max-width: 480px; margin: 0 auto; padding: 24px 16px; position: relative; z-index: 1; }

        /* ── レスポンシブ ── */
        @media (min-width: 768px) {
            .wrap { max-width: 760px; padding: 40px 32px; }
            .logo { font-size: 2.6rem; }
            .subtitle { font-size: 0.92rem; }
            .card { padding: 28px; }
            .form-grid { grid-template-columns: 1fr 1fr 2fr; }
            .form-full { grid-column: auto; }
            .room-grid { grid-template-columns: repeat(5, 1fr); }
        }
        @media (min-width: 1100px) {
            .wrap { max-width: 1100px; padding: 48px 40px; }
            .page-cols { display: grid; grid-template-columns: 320px 1fr; gap: 32px; align-items: start; }
            .sidebar { position: sticky; top: 28px; }
            .logo { font-size: 2.8rem; }
            .room-grid { grid-template-columns: repeat(7, 1fr); gap: 10px; }
            .modal { border-radius: 20px; max-width: 460px; margin: auto; }
            .modal-overlay { align-items: center; }
        }

        /* ── ヘッダー ── */
        header { margin-bottom: 28px; }
        .logo {
            font-family: 'Syne', sans-serif;
            font-size: 2rem; font-weight: 800;
            letter-spacing: -0.03em;
            background: linear-gradient(135deg, #fff 0%, var(--accent) 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .logo span { font-weight: 700; opacity: 0.5; font-size: 1.1rem; }
        .subtitle { color: var(--muted); font-size: 0.82rem; margin-top: 4px; letter-spacing: 0.03em; }

        /* ── 現在時刻バッジ ── */
        .now-badge {
            display: inline-flex; align-items: center; gap: 6px;
            background: var(--surface2); border: 1px solid var(--border);
            border-radius: 99px; padding: 5px 12px; font-size: 0.78rem;
            color: var(--muted); margin-top: 10px;
        }
        .now-badge .dot { width: 6px; height: 6px; border-radius: 50%; background: #4dd9a0; animation: pulse 2s infinite; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }

        /* ── 検索フォーム ── */
        .card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 20px;
            margin-bottom: 16px;
        }

        .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }
        .form-full { grid-column: 1 / -1; }

        label { display: block; font-size: 0.72rem; font-weight: 700; color: var(--muted); letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 6px; }

        select {
            width: 100%; padding: 12px 14px;
            background: var(--surface2); border: 1px solid var(--border);
            border-radius: 10px; color: var(--text); font-size: 0.95rem;
            font-family: 'Noto Sans JP', sans-serif;
            appearance: none;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%236b6b80' d='M6 8L1 3h10z'/%3E%3C/svg%3E");
            background-repeat: no-repeat; background-position: right 12px center;
            cursor: pointer; transition: border-color 0.2s;
        }
        select:focus { outline: none; border-color: var(--accent); }

        .btn-search {
            width: 100%; padding: 15px;
            background: var(--accent);
            border: none; border-radius: 10px;
            color: #fff; font-size: 1rem; font-weight: 700;
            font-family: 'Noto Sans JP', sans-serif;
            cursor: pointer; letter-spacing: 0.04em;
            transition: opacity 0.2s, transform 0.1s;
            position: relative; overflow: hidden;
        }
        .btn-search:active { transform: scale(0.98); }
        .btn-search:hover { opacity: 0.88; }

        /* ── 結果ヘッダー ── */
        .result-meta {
            display: flex; justify-content: space-between; align-items: center;
            margin-bottom: 14px;
        }
        .result-label { font-size: 0.78rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.08em; }
        .count-chip {
            font-family: 'Syne', sans-serif; font-size: 0.9rem; font-weight: 700;
            background: var(--surface2); border: 1px solid var(--border);
            padding: 4px 12px; border-radius: 99px; color: var(--text);
        }
        .count-chip em { color: var(--accent); font-style: normal; }

        /* ── 校舎セクション ── */
        .building-section { margin-bottom: 20px; }
        .building-label {
            display: flex; align-items: center; gap: 8px;
            font-size: 0.72rem; font-weight: 700; letter-spacing: 0.1em;
            text-transform: uppercase; margin-bottom: 10px;
            padding-bottom: 8px; border-bottom: 1px solid var(--border);
        }
        .building-label .badge {
            font-size: 0.68rem; padding: 2px 8px; border-radius: 99px; font-weight: 700;
        }
        .badge-tower    { background: var(--tower-bg);     color: var(--tower-color);     border: 1px solid var(--tower-border); }
        .badge-surugadai{ background: var(--surugadai-bg); color: var(--surugadai-color); border: 1px solid var(--surugadai-border); }
        .badge-funabashi{ background: var(--funabashi-bg); color: var(--funabashi-color); border: 1px solid var(--funabashi-border); }

        /* ── 教室カード ── */
        .room-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
        @media (max-width: 340px) { .room-grid { grid-template-columns: repeat(2, 1fr); } }

        .room-card {
            border-radius: 10px; padding: 12px 8px;
            text-align: center; cursor: pointer;
            transition: transform 0.15s, opacity 0.15s;
            border: 1px solid transparent;
            position: relative;
        }
        .room-card:active { transform: scale(0.96); }

        .room-card.tower     { background: var(--tower-bg);     border-color: var(--tower-border); }
        .room-card.surugadai { background: var(--surugadai-bg); border-color: var(--surugadai-border); }
        .room-card.funabashi { background: var(--funabashi-bg); border-color: var(--funabashi-border); }

        .room-number {
            font-family: 'Syne', sans-serif; font-size: 1.15rem; font-weight: 700;
            display: block; line-height: 1.2;
        }
        .room-card.tower     .room-number { color: var(--tower-color); }
        .room-card.surugadai .room-number { color: var(--surugadai-color); }
        .room-card.funabashi .room-number { color: var(--funabashi-color); }

        .room-bldg { font-size: 0.65rem; color: var(--muted); margin-top: 3px; display: block; }

        /* 予約中バッジ */
        .room-card.reserved::after {
            content: '予約中';
            position: absolute; top: 5px; right: 5px;
            font-size: 0.55rem; font-weight: 700;
            background: rgba(255,255,255,0.1); color: var(--muted);
            padding: 1px 5px; border-radius: 4px;
        }

        /* ── モーダル（仮予約） ── */
        .modal-overlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,0.7); backdrop-filter: blur(6px);
            z-index: 100; align-items: flex-end; justify-content: center;
        }
        .modal-overlay.open { display: flex; }

        .modal {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 20px 20px 0 0;
            padding: 28px 24px 36px;
            width: 100%; max-width: 480px;
            animation: slideUp 0.25s cubic-bezier(0.34, 1.56, 0.64, 1);
        }
        @keyframes slideUp { from { transform: translateY(100%); } to { transform: translateY(0); } }

        .modal-handle { width: 40px; height: 4px; background: var(--border); border-radius: 2px; margin: 0 auto 20px; }
        .modal-title { font-family: 'Syne', sans-serif; font-size: 1.3rem; font-weight: 800; margin-bottom: 6px; }
        .modal-room-info { color: var(--muted); font-size: 0.85rem; margin-bottom: 20px; }

        .modal-notice {
            background: rgba(255,200,60,0.08); border: 1px solid rgba(255,200,60,0.2);
            border-radius: 10px; padding: 12px 14px; font-size: 0.78rem;
            color: #ffc83c; line-height: 1.6; margin-bottom: 20px;
        }
        .modal-notice strong { display: block; margin-bottom: 2px; }

        .modal-form label { color: var(--muted); font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; display: block; margin-bottom: 6px; margin-top: 14px; }
        .modal-form input, .modal-form textarea {
            width: 100%; padding: 12px 14px;
            background: var(--surface2); border: 1px solid var(--border);
            border-radius: 10px; color: var(--text); font-size: 0.9rem;
            font-family: 'Noto Sans JP', sans-serif;
        }
        .modal-form input:focus, .modal-form textarea:focus { outline: none; border-color: var(--accent); }
        .modal-form textarea { resize: none; height: 72px; }

        .modal-actions { display: grid; grid-template-columns: 1fr 2fr; gap: 10px; margin-top: 18px; }
        .btn-cancel {
            padding: 13px; background: var(--surface2); border: 1px solid var(--border);
            border-radius: 10px; color: var(--muted); font-size: 0.9rem; font-weight: 700;
            font-family: 'Noto Sans JP', sans-serif; cursor: pointer;
        }
        .btn-reserve {
            padding: 13px; background: var(--accent); border: none;
            border-radius: 10px; color: #fff; font-size: 0.9rem; font-weight: 700;
            font-family: 'Noto Sans JP', sans-serif; cursor: pointer;
        }

        /* ── トースト ── */
        .toast {
            display: none; position: fixed; bottom: 32px; left: 50%; transform: translateX(-50%);
            background: var(--surface); border: 1px solid var(--border);
            border-radius: 99px; padding: 10px 20px; font-size: 0.85rem;
            color: var(--text); z-index: 200; white-space: nowrap;
            box-shadow: 0 8px 32px rgba(0,0,0,0.4);
        }
        .toast.show { display: block; animation: fadeInOut 2.5s ease forwards; }
        @keyframes fadeInOut {
            0%{opacity:0;transform:translateX(-50%) translateY(10px)}
            15%{opacity:1;transform:translateX(-50%) translateY(0)}
            75%{opacity:1}
            100%{opacity:0}
        }

        /* ── 空き無し ── */
        .empty-state { text-align: center; padding: 40px 20px; color: var(--muted); font-size: 0.9rem; }
        .empty-state .icon { font-size: 2.5rem; display: block; margin-bottom: 10px; }

        /* ── エラー ── */
        .error-card { background: rgba(255,80,80,0.08); border: 1px solid rgba(255,80,80,0.2); border-radius: var(--radius); padding: 16px; color: #ff6060; font-size: 0.85rem; }
    </style>
</head>
<body>
<div class="wrap">

    <!-- ヘッダー -->
    <header>
        <div class="logo">RoomRadar <span>β</span></div>
        <div class="subtitle">日大理工学部 · 空き教室リアルタイム検索</div>
        <div class="now-badge">
            <span class="dot"></span>
            <span id="clock">--:--</span>
            <span>·</span>
            <span id="now-period">現在時限 検出中…</span>
        </div>
    </header>

    <!-- PC: 2カラム / スマホ・タブレット: 1カラム -->
    <div class="page-cols">

        <!-- 左サイドバー（PC）/ 上部（スマホ） -->
        <div class="sidebar">
            <div class="card">
                <form method="POST" id="search-form">
                    <div class="form-grid">
                        <div>
                            <label>曜日</label>
                            <select name="day" id="sel-day">
                                {% for d in ["月", "火", "水", "木", "金", "土"] %}
                                <option value="{{ d }}" {% if selected_day == d %}selected{% endif %}>{{ d }}曜日</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div>
                            <label>時限</label>
                            <select name="period" id="sel-period">
                                {% for p in range(1, 7) %}
                                <option value="{{ p }}" {% if selected_period == p %}selected{% endif %}>{{ p }}限 ({{ period_times[p] }})</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="form-full">
                            <label>校舎</label>
                            <select name="building">
                                <option value="all">すべての校舎</option>
                                <option value="tower"     {% if selected_building == 'tower'     %}selected{% endif %}>🏢 タワースコラ</option>
                                <option value="main"      {% if selected_building == 'main'      %}selected{% endif %}>🏫 駿河台校舎</option>
                                <option value="funabashi" {% if selected_building == 'funabashi' %}selected{% endif %}>🏛 船橋校舎</option>
                            </select>
                        </div>
                    </div>
                    <button type="submit" class="btn-search">空き教室を検索</button>
                </form>
            </div>
        </div><!-- /sidebar -->

        <!-- 右: 結果エリア（PC）/ 下部（スマホ） -->
        <div class="results-area">

            <!-- エラー -->
            {% if error_message %}
            <div class="error-card">⚠ {{ error_message }}</div>
            {% endif %}

            <!-- 結果 -->
            {% if empty_rooms is not none and not error_message %}
            <div class="result-meta">
                <span class="result-label">{{ selected_day }}曜 {{ selected_period }}限 の空き教室</span>
                <span class="count-chip"><em>{{ empty_rooms|length }}</em> 室</span>
            </div>

            {% if empty_rooms %}
                {% set tower_rooms     = empty_rooms | selectattr('building', 'equalto', 'タワースコラ') | list %}
                {% set surugadai_rooms = empty_rooms | selectattr('building', 'equalto', '駿河台校舎') | list %}
                {% set funabashi_rooms = empty_rooms | selectattr('building', 'equalto', '船橋校舎') | list %}

                {% if tower_rooms %}
                <div class="building-section">
                    <div class="building-label">
                        <span class="badge badge-tower">タワースコラ</span>
                        {{ tower_rooms|length }}室
                    </div>
                    <div class="room-grid">
                    {% for room in tower_rooms %}
                        <div class="room-card tower" onclick="openModal('{{ room.name }}', 'タワースコラ')">
                            <span class="room-number">{{ room.name }}</span>
                            <span class="room-bldg">タワースコラ</span>
                        </div>
                    {% endfor %}
                    </div>
                </div>
                {% endif %}

                {% if surugadai_rooms %}
                <div class="building-section">
                    <div class="building-label">
                        <span class="badge badge-surugadai">駿河台校舎</span>
                        {{ surugadai_rooms|length }}室
                    </div>
                    <div class="room-grid">
                    {% for room in surugadai_rooms %}
                        <div class="room-card surugadai" onclick="openModal('{{ room.name }}', '駿河台校舎')">
                            <span class="room-number">{{ room.name }}</span>
                            <span class="room-bldg">駿河台校舎</span>
                        </div>
                    {% endfor %}
                    </div>
                </div>
                {% endif %}

                {% if funabashi_rooms %}
                <div class="building-section">
                    <div class="building-label">
                        <span class="badge badge-funabashi">船橋校舎</span>
                        {{ funabashi_rooms|length }}室
                    </div>
                    <div class="room-grid">
                    {% for room in funabashi_rooms %}
                        <div class="room-card funabashi" onclick="openModal('{{ room.name }}', '船橋校舎')">
                            <span class="room-number">{{ room.name }}</span>
                            <span class="room-bldg">船橋校舎</span>
                        </div>
                    {% endfor %}
                    </div>
                </div>
                {% endif %}

            {% else %}
            <div class="empty-state">
                <span class="icon">🔍</span>
                条件に合う空き教室はありませんでした。
            </div>
            {% endif %}
            {% endif %}

        </div><!-- /results-area -->
    </div><!-- /page-cols -->

</div><!-- /wrap -->

<!-- 仮予約モーダル -->
<div class="modal-overlay" id="modal" onclick="handleOverlayClick(event)">
    <div class="modal">
        <div class="modal-handle"></div>
        <div class="modal-title" id="modal-title">教室名</div>
        <div class="modal-room-info" id="modal-info">校舎 · 現在検索中の時間帯</div>

        <div class="modal-notice">
            <strong>⚠ これはRoomRadar上の仮予約です</strong>
            大学公式の予約システムではありません。教室の使用権を保証するものではなく、
            他の利用者への情報共有を目的としています。
        </div>

        <div class="modal-form">
            <label>お名前 / グループ名</label>
            <input type="text" id="reserve-name" placeholder="例: 二川・物理学科" maxlength="30">
            <label>利用目的（任意）</label>
            <textarea id="reserve-note" placeholder="例: グループ課題の打ち合わせ"></textarea>
        </div>

        <div class="modal-actions">
            <button class="btn-cancel" onclick="closeModal()">キャンセル</button>
            <button class="btn-reserve" onclick="submitReserve()">仮予約する</button>
        </div>
    </div>
</div>

<!-- トースト -->
<div class="toast" id="toast"></div>

<script>
// ── 時計 & 現在時限の自動検出 ──
const PERIODS = {
    1: ["09:00","10:30"], 2: ["10:40","12:10"], 3: ["13:00","14:30"],
    4: ["14:40","16:10"], 5: ["16:20","17:50"], 6: ["18:00","19:30"]
};
const DAYS = ["日","月","火","水","木","金","土"];

function toMin(t) { const [h,m] = t.split(':').map(Number); return h*60+m; }

function updateClock() {
    const now = new Date();
    const jst = new Date(now.toLocaleString('en', {timeZone:'Asia/Tokyo'}));
    const h = String(jst.getHours()).padStart(2,'0');
    const m = String(jst.getMinutes()).padStart(2,'0');
    document.getElementById('clock').textContent = h + ':' + m;

    const cur = h*60 + jst.getMinutes();
    let found = null;
    for (const [p, [s,e]] of Object.entries(PERIODS)) {
        if (cur >= toMin(s) && cur <= toMin(e)) { found = p; break; }
    }
    document.getElementById('now-period').textContent = found ? found + '限 授業中' : '授業時間外';
}
setInterval(updateClock, 1000);
updateClock();

// ── ページ読み込み時に現在の曜日・時限をセット（GETのみ） ──
{% if not searched %}
(function() {
    const now = new Date();
    const jst = new Date(now.toLocaleString('en', {timeZone:'Asia/Tokyo'}));
    const dayIdx = jst.getDay(); // 0=日
    const dayName = DAYS[dayIdx === 0 ? 1 : dayIdx]; // 日曜は月に
    const sel = document.getElementById('sel-day');
    for (let i = 0; i < sel.options.length; i++) {
        if (sel.options[i].value === dayName) { sel.selectedIndex = i; break; }
    }
    const cur = jst.getHours()*60 + jst.getMinutes();
    let period = 1;
    for (const [p, [s,e]] of Object.entries(PERIODS)) {
        if (cur >= toMin(s) && cur <= toMin(e)) { period = parseInt(p); break; }
        if (cur < toMin(s)) { period = Math.max(1, parseInt(p)-1); break; }
    }
    document.getElementById('sel-period').value = period;
})();
{% endif %}

// ── 仮予約モーダル ──
let currentRoom = '', currentBuilding = '';

function openModal(room, building) {
    currentRoom = room;
    currentBuilding = building;
    document.getElementById('modal-title').textContent = room;
    document.getElementById('modal-info').textContent =
        building + ' · {{ selected_day }}曜 {{ selected_period }}限';
    document.getElementById('reserve-name').value = '';
    document.getElementById('reserve-note').value = '';
    document.getElementById('modal').classList.add('open');
}

function closeModal() {
    document.getElementById('modal').classList.remove('open');
}

function handleOverlayClick(e) {
    if (e.target === document.getElementById('modal')) closeModal();
}

function submitReserve() {
    const name = document.getElementById('reserve-name').value.trim();
    if (!name) {
        showToast('⚠ お名前を入力してください');
        return;
    }
    closeModal();
    showToast('✓ ' + currentRoom + ' を仮予約しました（RoomRadar上のみ）');
}

function showToast(msg) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.remove('show');
    void t.offsetWidth;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 2600);
}
</script>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def index():
    now = datetime.datetime.now(JST)
    day = ["月", "火", "水", "木", "金", "土", "日"][now.weekday()]
    if day == "日": day = "月"

    c_time = now.strftime("%H:%M")
    period = 1
    for p, (s, e) in PERIODS.items():
        if s <= c_time <= e:
            period = p
            break

    building = "all"
    empty_rooms = None
    error_message = None
    searched = False

    period_times = {p: f"{s}–{e}" for p, (s, e) in PERIODS.items()}

    if request.method == 'POST':
        searched = True
        day = request.form.get('day')
        period = int(request.form.get('period'))
        building = request.form.get('building')

    try:
        if not os.path.exists(DB_NAME):
            raise FileNotFoundError(f"データベースファイル '{DB_NAME}' が見つかりません。")

        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()

        placeholders = ','.join(['?'] * len(ACTIVE_TERMS))
        cur.execute(
            f"SELECT 教室 FROM schedules WHERE 曜日=? AND 時限=? AND 履修期名 IN ({placeholders})",
            [day, period] + ACTIVE_TERMS
        )
        occupied = {str(row[0]) for row in cur.fetchall()}

        q_all = "SELECT name, building FROM classrooms"
        if building == "tower":
            q_all += " WHERE building = 'タワースコラ'"
        elif building == "main":
            q_all += " WHERE building = '駿河台校舎'"
        elif building == "funabashi":
            q_all += " WHERE building = '船橋校舎'"

        cur.execute(q_all)
        all_rooms = cur.fetchall()
        conn.close()

        empty_rooms = sorted(
            [{"name": r[0], "building": r[1]} for r in all_rooms if str(r[0]) not in occupied],
            key=lambda x: (x['building'] != 'タワースコラ', x['building'] != '駿河台校舎', x['name'])
        )

    except Exception as e:
        error_message = f"システムエラーが発生しました: {e}"
        empty_rooms = []

    return render_template_string(
        HTML_TEMPLATE,
        empty_rooms=empty_rooms,
        selected_day=day, selected_period=period,
        selected_building=building,
        error_message=error_message,
        period_times=period_times,
        searched=searched
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
