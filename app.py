
import os
import glob
import sqlite3
import datetime
import pandas as pd
from flask import Flask, render_template_string, request

app = Flask(__name__)

# データ読み込み・DB構築ロジック
DB_NAME = "schedule_final.db"

def init_db():
    print("🔄 データベースを初期化中...")
    xlsx_files = glob.glob('*.xlsx')
    if not xlsx_files:
        print("⚠️ データファイルがありません")
        return False
    
    filename = xlsx_files[0]
    
    try:
        xls = pd.ExcelFile(filename)
        df_stats = None
        df_schedule = None
        
        # シートの自動判別
        for sheet in xls.sheet_names:
            df = pd.read_excel(filename, sheet_name=sheet)
            cols = str(df.columns)
            if '曜日' in cols and '時限' in cols:
                df_schedule = df
            elif 'Class_Count' in cols or 'Classroom_Clean' in cols:
                df_stats = df
                
        if df_stats is None or df_schedule is None:
            return False

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('DROP TABLE IF EXISTS schedules')
        cursor.execute('DROP TABLE IF EXISTS classrooms')
        
        # ---------------------------------------------------------
        # 【修正箇所】校舎の振り分けロジック改良
        # ---------------------------------------------------------
        cursor.execute('''CREATE TABLE classrooms (id INTEGER PRIMARY KEY, name TEXT, building TEXT, capacity INTEGER)''')
        classroom_data = []
        
        # カラム名の特定
        name_col = 'Classroom_Clean' if 'Classroom_Clean' in df_stats.columns else df_stats.columns[0]
        campus_col = '設置校舎' if '設置校舎' in df_stats.columns else None

        for _, row in df_stats.iterrows():
            room_name = str(row[name_col])
            
            # 元データの「設置校舎」情報を取得（もしあれば）
            campus_raw = str(row[campus_col]) if campus_col and campus_col in row else ""
            
            # --- 新しい振り分けロジック ---
            if "船橋" in campus_raw:
                building = "船橋校舎"
            elif room_name.isdigit() and len(room_name) == 4: # データがない場合のバックアップ(4桁数字)
                building = "船橋校舎"
            elif room_name.startswith('S'):
                building = "タワースコラ"
            else:
                building = "駿河台校舎"
            # ---------------------------

            capacity = row['Class_Count'] if 'Class_Count' in row else 0
            classroom_data.append((room_name, building, capacity))
            
        cursor.executemany('INSERT INTO classrooms (name, building, capacity) VALUES (?, ?, ?)', classroom_data)
        
        # スケジュールデータの登録
        cursor.execute('''CREATE TABLE schedules (id INTEGER PRIMARY KEY, day TEXT, period INTEGER, classroom_name TEXT, term TEXT, subject_code TEXT)''')
        schedule_data = []
        room_col_sch = 'Classroom_Clean' if 'Classroom_Clean' in df_schedule.columns else '教室'
        
        for _, row in df_schedule.iterrows():
            if room_col_sch not in row: continue
            schedule_data.append((row['曜日'], row['時限'], str(row[room_col_sch]), row.get('履修期名', '通年'), row.get('時間割CD', '')))
            
        cursor.executemany('INSERT INTO schedules (day, period, classroom_name, term, subject_code) VALUES (?, ?, ?, ?, ?)', schedule_data)
        conn.commit()
        conn.close()
        print("✅ データベース構築完了 (船橋校舎対応済み)")
        return True
    except Exception as e:
        print(f"❌ DB Error: {e}")
        return False

# 起動時にDBを作成
init_db()

JST = datetime.timezone(datetime.timedelta(hours=9))
PERIODS = {1: ("09:00", "10:30"), 2: ("10:40", "12:10"), 3: ("13:00", "14:30"),
           4: ("14:40", "16:10"), 5: ("16:20", "17:50"), 6: ("18:00", "19:30")}
ACTIVE_TERMS = ['後期', '年間', '後期隔週', '年間隔週', '後期集中(資)', '年間集中(資)', '年隔集中(資)']

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>理工学部 空き教室検索</title>
    <style>
        body { font-family: -apple-system, sans-serif; padding: 20px; background: #f8f9fa; color: #333; max-width: 600px; margin: 0 auto; }
        h1 { text-align: center; color: #003366; font-size: 1.5rem; }
        .card { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin-bottom: 20px; }
        select, button { width: 100%; padding: 12px; margin: 8px 0; border-radius: 8px; border: 1px solid #ddd; font-size: 16px; }
        button { background: #0056b3; color: white; border: none; font-weight: bold; cursor: pointer; }
        .result-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .count-badge { background: #28a745; color: white; padding: 4px 10px; border-radius: 20px; font-size: 0.9rem; font-weight: bold; }
        .room-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 10px; }
        .room-item { background: white; border: 1px solid #eee; padding: 10px; border-radius: 8px; text-align: center; }
        .room-name { font-size: 1.2rem; font-weight: bold; color: #333; display: block; }
        .room-info { font-size: 0.75rem; color: #888; margin-top: 4px; }
        
        /* 校舎ごとの色分け */
        .tower { border-left: 5px solid #007bff; background: #f0f7ff; } /* 青 */
        .main { border-left: 5px solid #28a745; background: #f0fff4; }  /* 緑 */
        .funabashi { border-left: 5px solid #fd7e14; background: #fff5f0; } /* オレンジ */
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
                <option value="tower" {% if selected_building == 'tower' %}selected{% endif %}>タワースコラ (S棟)</option>
                <option value="main" {% if selected_building == 'main' %}selected{% endif %}>駿河台校舎 (1号館等)</option>
                <option value="funabashi" {% if selected_building == 'funabashi' %}selected{% endif %}>船橋校舎</option>
            </select>
            <button type="submit">検索</button>
        </form>
    </div>

    {% if empty_rooms is not none %}
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
        <div style="grid-column: 1/-1; text-align:center; padding:20px; color:#888;">条件に合う空き教室はありません...</div>
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
    c_time = now.strftime("%H:%M")
    period = 1
    for p, (s, e) in PERIODS.items():
        if s <= c_time <= e: period = p; break
            
    building = "all"
    empty_rooms = None
    
    if request.method == 'POST':
        day = request.form.get('day')
        period = int(request.form.get('period'))
        building = request.form.get('building')
    
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    placeholders = ','.join(['?'] * len(ACTIVE_TERMS))
    cur.execute(f"SELECT classroom_name FROM schedules WHERE day=? AND period=? AND term IN ({placeholders})", [day, period] + ACTIVE_TERMS)
    occupied = {str(row[0]) for row in cur.fetchall()}
    
    # フィルタリング条件の修正
    q_all = "SELECT name, building FROM classrooms"
    params = []
    
    if building == "tower": 
        q_all += " WHERE building = 'タワースコラ'"
    elif building == "main": 
        q_all += " WHERE building = '駿河台校舎'"
    elif building == "funabashi":
        q_all += " WHERE building = '船橋校舎'"
        
    cur.execute(q_all, params)
    all_rooms = cur.fetchall()
    conn.close()
    
    # ソート順: タワー -> 駿河台 -> 船橋
    empty_rooms = sorted([{"name": r[0], "building": r[1]} for r in all_rooms if str(r[0]) not in occupied], 
                         key=lambda x: (x['building'] != 'タワースコラ', x['building'] != '駿河台校舎', x['name']))
                         
    return render_template_string(HTML_TEMPLATE, empty_rooms=empty_rooms, selected_day=day, selected_period=period, selected_building=building)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
