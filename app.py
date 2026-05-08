import os
import sqlite3
import datetime
from flask import Flask, render_template_string, request

app = Flask(__name__)

# ==========================================
# RoomRadar 2026 検索エンジン・コア
# ==========================================
# 事前構築済みの完全版データベースを使用
DB_NAME = "schedule_final.db"

JST = datetime.timezone(datetime.timedelta(hours=9))
PERIODS = {1: ("09:00", "10:30"), 2: ("10:40", "12:10"), 3: ("13:00", "14:30"),
           4: ("14:40", "16:10"), 5: ("16:20", "17:50"), 6: ("18:00", "19:30")}

# 現在のデータベース（前期版）に合わせた検索対象ターム
ACTIVE_TERMS = ['前期', '通年']

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>理工学部 空き教室検索 (RoomRadar)</title>
    <style>
        body { font-family: -apple-system, sans-serif; padding: 20px; background: #f8f9fa; color: #333; max-width: 600px; margin: 0 auto; }
        h1 { text-align: center; color: #003366; font-size: 1.5rem; }
        .card { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin-bottom: 20px; }
        select, button { width: 100%; padding: 12px; margin: 8px 0; border-radius: 8px; border: 1px solid #ddd; font-size: 16px; }
        button { background: #0056b3; color: white; border: none; font-weight: bold; cursor: pointer; transition: background 0.2s; }
        button:hover { background: #004494; }
        .result-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .count-badge { background: #28a745; color: white; padding: 4px 10px; border-radius: 20px; font-size: 0.9rem; font-weight: bold; }
        .room-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 10px; }
        .room-item { background: white; border: 1px solid #eee; padding: 10px; border-radius: 8px; text-align: center; }
        .room-name { font-size: 1.2rem; font-weight: bold; color: #333; display: block; }
        .room-info { font-size: 0.75rem; color: #888; margin-top: 4px; }
        
        /* 校舎ごとの色分け */
        .tower { border-left: 5px solid #007bff; background: #f0f7ff; } 
        .main { border-left: 5px solid #28a745; background: #f0fff4; }  
        .funabashi { border-left: 5px solid #fd7e14; background: #fff5f0; } 
        .error-msg { color: #dc3545; font-weight: bold; text-align: center; padding: 10px; }
    </style>
</head>
<body>
    <h1>理工学部 空き教室検索</h1>
    <div class="card">
        <form method="POST">
            <select name="day">
                {% for d in ["月", "火", "水", "木", "金", "土"] %}
                <option value="{{ d }}" {% if selected_day == d %}selected{% endif %}>{{ d }}曜日</option>
                {% endfor %}
            </select>
            <select name="period">
                {% for p in range(1, 7) %}
                <option value="{{ p }}" {% if selected_period == p %}selected{% endif %}>{{ p }}限</option>
                {% endfor %}
            </select>
            <select name="building">
                <option value="all">すべての校舎</option>
                <option value="tower" {% if selected_building == 'tower' %}selected{% endif %}>タワースコラ</option>
                <option value="main" {% if selected_building == 'main' %}selected{% endif %}>駿河台校舎</option>
                <option value="funabashi" {% if selected_building == 'funabashi' %}selected{% endif %}>船橋校舎</option>
            </select>
            <button type="submit">空き教室を検索</button>
        </form>
    </div>

    {% if error_message %}
    <div class="card error-msg">
        {{ error_message }}
    </div>
    {% endif %}

    {% if empty_rooms is not none and not error_message %}
    <div class="result-header">
        <strong>検索結果</strong>
        <span class="count-badge">{{ empty_rooms|length }} 教室 空き</span>
    </div>
    <div class="room-list">
        {% for room in empty_rooms %}
        <div class="room-item {% if 'タワー' in room.building %}tower{% elif '船橋' in room.building %}funabashi{% else %}main{% endif %}">
            <span class="room-name">{{ room.name }}</span>
            <div class="room-info">{{ room.building }}</div>
        </div>
        {% else %}
        <div style="grid-column: 1/-1; text-align:center; padding:20px; color:#888;">条件に合う空き教室はありません。</div>
        {% endfor %}
    </div>
    {% endif %}
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def index():
    now = datetime.datetime.now(JST)
    day = ["月", "火", "水", "木", "金", "土", "日"][now.weekday()]
    
    # 日曜日の場合はデフォルトを月曜にする
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
    
    if request.method == 'POST':
        day = request.form.get('day')
        period = int(request.form.get('period'))
        building = request.form.get('building')
    
    try:
        if not os.path.exists(DB_NAME):
            raise FileNotFoundError(f"データベースファイル '{DB_NAME}' が見つかりません。")

        conn = sqlite3.connect(DB_NAME)
        cur = conn.cursor()
        
        # 1. 指定した曜日・時限・履修期において「使用中」の教室名を抽出
        placeholders = ','.join(['?'] * len(ACTIVE_TERMS))
        query_occupied = f"SELECT 教室 FROM schedules WHERE 曜日=? AND 時限=? AND 履修期名 IN ({placeholders})"
        cur.execute(query_occupied, [day, period] + ACTIVE_TERMS)
        occupied = {str(row[0]) for row in cur.fetchall()}
        
        # 2. 全教室マスタから対象キャンパスの教室を抽出
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
        
        # 3. 差分抽出（全教室 - 使用中教室 = 空き教室）およびソート
        empty_rooms = sorted([{"name": r[0], "building": r[1]} for r in all_rooms if str(r[0]) not in occupied], 
                             key=lambda x: (x['building'] != 'タワースコラ', x['building'] != '駿河台校舎', x['name']))
                             
    except Exception as e:
        error_message = f"システムエラーが発生しました: {e}"
        empty_rooms = []
                         
    return render_template_string(HTML_TEMPLATE, empty_rooms=empty_rooms, 
                                  selected_day=day, selected_period=period, 
                                  selected_building=building, error_message=error_message)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
