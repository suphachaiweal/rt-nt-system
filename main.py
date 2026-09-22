import pandas as pd
import numpy as np
from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import List, Dict
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os
import shutil
from datetime import datetime, timedelta, timezone
import json
import psycopg2
from supabase import create_client, Client

app = FastAPI()

# ==========================================
SUPABASE_URL = "https://svrtsuffmpoaeaozuihv.supabase.co" 
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN2cnRzdWZmbXBvYWVhb3p1aWh2Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MDA0MDYwNiwiZXhwIjoyMTA1NjE2NjA2fQ.FCVn1kORiSFBafWeAEeC1bOz5-5ioPbrXvaCVfL6uQM"
DB_URL = "postgresql://postgres.svrtsuffmpoaeaozuihv:DlaAdmin2569Supabase@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres"
# ==========================================

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
os.makedirs("uploads", exist_ok=True)
app.mount("/uploads", StaticFiles(directory="uploads"), name="uploads")

try:
    df_excel = pd.read_excel('รายชื่อสถานศึกษา.xls', skiprows=2)
    df_excel.columns = ['No', 'Province', 'Amphoe', 'DLA', 'Sch_No', 'School']
    df_excel['Province'] = df_excel['Province'].ffill()
    df_excel['DLA'] = df_excel['DLA'].ffill()
    df_clean = df_excel.dropna(subset=['School']).copy()

    DB_DATA = {}
    for index, row in df_clean.iterrows():
        prov = str(row['Province']).strip()
        dla = str(row['DLA']).strip()
        sch = str(row['School']).strip()
        
        if prov not in DB_DATA:
            DB_DATA[prov] = {}
        if dla not in DB_DATA[prov]:
            DB_DATA[prov][dla] = []
        if sch not in DB_DATA[prov][dla]:
            DB_DATA[prov][dla].append(sch)
            
    PROVINCES = list(DB_DATA.keys())
except Exception as e:
    DB_DATA = {}
    PROVINCES = []

@app.get("/api/school_data")
async def get_school_data():
    return JSONResponse(content=DB_DATA)

def get_db_connection():
    return psycopg2.connect(DB_URL)

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS school_data
                 (id SERIAL PRIMARY KEY,
                  province TEXT,
                  dla_name TEXT,
                  school_name TEXT,
                  rt_student_count INTEGER,
                  nt_student_count INTEGER)''')
                  
    c.execute("ALTER TABLE school_data ADD COLUMN IF NOT EXISTS rt_special_json JSONB DEFAULT '{}'::jsonb")
    c.execute("ALTER TABLE school_data ADD COLUMN IF NOT EXISTS nt_special_json JSONB DEFAULT '{}'::jsonb")
                  
    c.execute('''CREATE TABLE IF NOT EXISTS province_uploads
                 (province TEXT PRIMARY KEY,
                  file_path TEXT)''')
                  
    c.execute('''CREATE TABLE IF NOT EXISTS province_locks
                 (province TEXT PRIMARY KEY,
                  ref_code TEXT,
                  locked_at TEXT)''')
    conn.commit()
    conn.close()

try: init_db()
except: pass

class School(BaseModel):
    school_name: str
    rt_count: int
    nt_count: int
    has_special: bool
    rt_special: Dict[str, int]
    nt_special: Dict[str, int]

class DLA(BaseModel):
    dla_name: str
    schools: List[School]

class Submission(BaseModel):
    province: str
    dlas: List[DLA]

def is_province_locked(province: str) -> bool:
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT province FROM province_locks WHERE province=%s", (province,))
    locked = c.fetchone() is not None
    conn.close()
    return locked

def get_province_options(selected=""):
    options = '<option value="">-- เลือกจังหวัด --</option>'
    for p in PROVINCES:
        sel = 'selected' if p == selected else ''
        options += f'<option value="{p}" {sel}>{p}</option>'
    return options

SPECIAL_CATEGORIES = [
    ("t1", "ทางการเห็น"), ("t2", "ทางการได้ยิน"), ("t3", "สติปัญญา"),
    ("t4", "ร่างกาย"), ("t5", "เรียนรู้ (LD)"), ("t6", "พูด/ภาษา"),
    ("t7", "พฤติกรรม"), ("t8", "ออทิสติก"), ("t9", "พิการซ้อน")
]

@app.get("/", response_class=HTMLResponse)
async def get_form():
    prov_opts = get_province_options()
    special_inputs_rt = "".join([f'<div class="sp-row"><label>{name}</label><input type="number" class="sp-input rt-sp-{cid}" min="0" value="0" oninput="calcSpecial(this, \'rt\')"></div>' for cid, name in SPECIAL_CATEGORIES])
    special_inputs_nt = "".join([f'<div class="sp-row"><label>{name}</label><input type="number" class="sp-input nt-sp-{cid}" min="0" value="0" oninput="calcSpecial(this, \'nt\')"></div>' for cid, name in SPECIAL_CATEGORIES])
    
    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>ระบบส่งข้อมูลนักเรียนสอบ RT/NT ประจำปีการศึกษา 2569</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: #f0f2f5; padding: 20px; color: #1a202c; }}
            .container {{ max-width: 900px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
            h2 {{ text-align: center; color: #2d3748; margin-bottom: 25px; line-height: 1.4; }}
            .header-section {{ margin-bottom: 20px; }}
            .dla-block {{ background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-bottom: 20px; position: relative; }}
            .dla-title {{ font-weight: 600; font-size: 16px; margin-bottom: 15px; color: #2b6cb0; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; }}
            .school-wrapper {{ background: white; padding: 15px; border-radius: 8px; border: 1px dashed #cbd5e0; margin-bottom: 15px; }}
            .school-block {{ display: flex; gap: 15px; align-items: flex-end; }}
            .school-block > div {{ flex: 1; }}
            .school-block .col-school {{ flex: 2; }}
            label {{ font-weight: 500; font-size: 14px; margin-bottom: 5px; display: block; color: #4a5568; }}
            input, select {{ width: 100%; padding: 8px 12px; border: 1px solid #cbd5e0; border-radius: 4px; box-sizing: border-box; font-family: 'Sarabun', sans-serif; }}
            input:focus, select:focus {{ outline: none; border-color: #2b6cb0; box-shadow: 0 0 0 1px #2b6cb0; }}
            .special-toggle {{ margin-top: 15px; background: #ebf8ff; padding: 8px 12px; border-radius: 4px; display: inline-block; cursor: pointer; }}
            .special-toggle input {{ width: auto; display: inline; margin-right: 8px; transform: scale(1.2); }}
            .special-panel {{ display: none; margin-top: 15px; padding-top: 15px; border-top: 1px dashed #cbd5e0; }}
            .sp-grid {{ display: flex; gap: 20px; }}
            .sp-col {{ flex: 1; background: #fdfaf4; padding: 15px; border-radius: 8px; border: 1px solid #f6e05e; }}
            .sp-col-title {{ font-weight: 600; text-align: center; color: #b7791f; margin-bottom: 10px; border-bottom: 1px solid #fbd38d; padding-bottom: 5px; }}
            .sp-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; }}
            .sp-row label {{ font-size: 13px; font-weight: 400; margin-bottom: 0; flex: 2; }}
            .sp-row input {{ flex: 1; padding: 4px 8px; text-align: center; }}
            .sp-summary {{ margin-top: 10px; padding: 8px; background: white; border-radius: 4px; text-align: center; font-size: 14px; font-weight: 600; box-shadow: inset 0 1px 3px rgba(0,0,0,0.1); }}
            .error-text {{ color: #e53e3e; font-size: 13px; margin-top: 5px; display: none; font-weight: 500; }}
            .btn {{ padding: 10px 15px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-family: 'Sarabun', sans-serif; font-weight: 500; transition: background 0.2s; }}
            .btn-primary {{ background-color: #2b6cb0; color: white; width: 100%; font-size: 16px; padding: 12px; margin-top: 20px; }}
            .btn-secondary {{ background-color: #edf2f7; color: #4a5568; border: 1px solid #cbd5e0; }}
            .btn-add-school {{ background-color: #ebf8ff; color: #2b6cb0; border: 1px solid #bee3f8; margin-top: 10px; width: 100%; }}
            .btn-remove {{ background-color: #fff5f5; color: #c53030; border: 1px solid #fed7d7; padding: 8px 12px; }}
            .nav-link {{ display: block; text-align: center; margin-top: 20px; text-decoration: none; color: #dd6b20; font-weight: 500; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>ระบบส่งข้อมูลนักเรียนสอบ RT/NT ประจำปีการศึกษา 2569<br>ระดับจังหวัด (บันทึกรวดเดียว)</h2>
            <div class="header-section">
                <label>จังหวัด (ผู้รายงาน)</label>
                <select id="provinceInput" required>
                    {prov_opts}
                </select>
            </div>
            <div id="dlaContainer"></div>
            <button type="button" class="btn btn-secondary" onclick="addDLA()" style="width: 100%; margin-bottom: 20px;">+ เพิ่ม อปท. อื่นในจังหวัดนี้</button>
            <button type="button" class="btn btn-primary" onclick="submitData()">💾 บันทึกและส่งข้อมูล</button>
            <a href="/dashboard" class="nav-link">📊 ตรวจสอบ/ปริ้นเอกสาร/แนบไฟล์รับรอง</a>
        </div>

        <script>
            let dbData = {{}};
            let dlaCount = 0;
            
            async function fetchSchoolData() {{
                const res = await fetch('/api/school_data');
                dbData = await res.json();
            }}
            
            document.getElementById('provinceInput').addEventListener('change', function() {{
                const dlaContainer = document.getElementById('dlaContainer');
                if (dlaContainer.innerHTML !== '') {{
                    if(!confirm('การเปลี่ยนจังหวัดจะล้างข้อมูลที่กำลังกรอก ต้องการเปลี่ยนหรือไม่?')) {{ return; }}
                }}
                dlaContainer.innerHTML = '';
                dlaCount = 0;
                if (this.value) {{ addDLA(); }}
            }});

            function createSchoolHTML(dlaId) {{
                return `
                    <div class="school-wrapper">
                        <div class="school-block">
                            <div class="col-school">
                                <label>ชื่อโรงเรียน</label>
                                <select class="school-name" required><option value="">-- เลือกโรงเรียน --</option></select>
                            </div>
                            <div>
                                <label>ยอด นร. ป.1 (RT) รวม</label>
                                <input type="number" class="rt-count" min="0" value="0" oninput="calcSpecial(this, 'rt', true)" required>
                            </div>
                            <div>
                                <label>ยอด นร. ป.3 (NT) รวม</label>
                                <input type="number" class="nt-count" min="0" value="0" oninput="calcSpecial(this, 'nt', true)" required>
                            </div>
                            <div style="flex: 0.3; text-align: center; margin-bottom: 2px;">
                                <button type="button" class="btn btn-remove" onclick="this.closest('.school-wrapper').remove()" title="ลบโรงเรียนนี้">ลบ</button>
                            </div>
                        </div>
                        
                        <label class="special-toggle">
                            <input type="checkbox" class="has-special" onchange="toggleSpecialPanel(this)"> 
                            ♿ โรงเรียนนี้มีนักเรียนที่มีความต้องการจำเป็นพิเศษ (เรียนร่วม)
                        </label>
                        
                        <div class="special-panel">
                            <div class="sp-grid">
                                <div class="sp-col">
                                    <div class="sp-col-title">ระบุประเภทพิเศษ ป.1 (RT)</div>
                                    {special_inputs_rt}
                                    <div class="error-text rt-error">❌ ยอดเด็กพิเศษรวมกัน มากกว่ายอดทั้งหมด!</div>
                                    <div class="sp-summary rt-summary">จำนวนเด็กปกติ: <span class="normal-num">0</span> คน</div>
                                </div>
                                <div class="sp-col">
                                    <div class="sp-col-title">ระบุประเภทพิเศษ ป.3 (NT)</div>
                                    {special_inputs_nt}
                                    <div class="error-text nt-error">❌ ยอดเด็กพิเศษรวมกัน มากกว่ายอดทั้งหมด!</div>
                                    <div class="sp-summary nt-summary">จำนวนเด็กปกติ: <span class="normal-num">0</span> คน</div>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
            }}
            
            function toggleSpecialPanel(cb) {{
                const panel = cb.closest('.school-wrapper').querySelector('.special-panel');
                panel.style.display = cb.checked ? 'block' : 'none';
                if(!cb.checked) {{
                    panel.querySelectorAll('.sp-input').forEach(inp => inp.value = 0);
                    calcSpecial(cb, 'rt', true);
                    calcSpecial(cb, 'nt', true);
                }}
            }}

            function calcSpecial(el, type, isMainInput=false) {{
                const wrapper = el.closest('.school-wrapper');
                const totalInput = wrapper.querySelector(`.${{type}}-count`);
                const total = parseInt(totalInput.value) || 0;
                
                let spTotal = 0;
                wrapper.querySelectorAll(`.${{type}}-sp-t1, .${{type}}-sp-t2, .${{type}}-sp-t3, .${{type}}-sp-t4, .${{type}}-sp-t5, .${{type}}-sp-t6, .${{type}}-sp-t7, .${{type}}-sp-t8, .${{type}}-sp-t9`).forEach(inp => {{
                    spTotal += parseInt(inp.value) || 0;
                }});

                const normalCount = total - spTotal;
                const summaryBox = wrapper.querySelector(`.${{type}}-summary`);
                const errorBox = wrapper.querySelector(`.${{type}}-error`);
                
                if(normalCount < 0) {{
                    summaryBox.style.display = 'none';
                    errorBox.style.display = 'block';
                    totalInput.style.borderColor = '#e53e3e';
                }} else {{
                    summaryBox.style.display = 'block';
                    errorBox.style.display = 'none';
                    totalInput.style.borderColor = '#cbd5e0';
                    summaryBox.querySelector('.normal-num').innerText = normalCount;
                }}
            }}
            
            function addDLA() {{
                dlaCount++;
                const dlaId = `dla-${{dlaCount}}`;
                const prov = document.getElementById('provinceInput').value;
                let dlaOptions = '<option value="">-- เลือก อปท. --</option>';
                if (prov && dbData[prov]) {{
                    for (let d in dbData[prov]) {{
                        dlaOptions += `<option value="${{d}}">${{d}}</option>`;
                    }}
                }}
                
                const dlaHTML = `
                    <div class="dla-block" id="${{dlaId}}">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div class="dla-title">อปท. ที่ ${{dlaCount}}</div>
                            ${{dlaCount > 1 ? `<button type="button" class="btn btn-remove" style="padding: 4px 8px; font-size: 12px;" onclick="document.getElementById('${{dlaId}}').remove()">- ลบ อปท. นี้</button>` : ''}}
                        </div>
                        <label>ชื่อ อปท.</label>
                        <select class="dla-name" required style="margin-bottom: 15px;" onchange="updateSchools(this, '${{dlaId}}')">
                            ${{dlaOptions}}
                        </select>
                        <div class="schools-container" id="schools-${{dlaId}}"></div>
                        <button type="button" class="btn btn-add-school" onclick="addSchool('${{dlaId}}')">+ เพิ่มโรงเรียนใน อปท. นี้</button>
                    </div>
                `;
                document.getElementById('dlaContainer').insertAdjacentHTML('beforeend', dlaHTML);
                addSchool(dlaId);
            }}
            
            function addSchool(dlaId) {{
                const container = document.getElementById(`schools-${{dlaId}}`);
                container.insertAdjacentHTML('beforeend', createSchoolHTML(dlaId));
                const dlaSelect = document.querySelector(`#${{dlaId}} .dla-name`);
                if(dlaSelect && dlaSelect.value) {{
                    const newSchoolSelect = container.lastElementChild.querySelector('.school-name');
                    populateSchoolSelect(newSchoolSelect, dlaSelect.value);
                }}
            }}

            function updateSchools(dlaSelect, dlaId) {{
                const dlaName = dlaSelect.value;
                const schoolSelects = document.querySelectorAll(`#${{dlaId}} .school-name`);
                schoolSelects.forEach(select => {{ populateSchoolSelect(select, dlaName); }});
            }}
            
            function populateSchoolSelect(selectElement, dlaName) {{
                const prov = document.getElementById('provinceInput').value;
                const currentValue = selectElement.value;
                selectElement.innerHTML = '<option value="">-- เลือกโรงเรียน --</option>';
                if (prov && dlaName && dbData[prov] && dbData[prov][dlaName]) {{
                    dbData[prov][dlaName].forEach(sch => {{
                        const opt = document.createElement('option');
                        opt.value = sch;
                        opt.text = sch;
                        selectElement.appendChild(opt);
                    }});
                    if(currentValue && dbData[prov][dlaName].includes(currentValue)) {{
                        selectElement.value = currentValue;
                    }}
                }}
            }}
            
            async function submitData() {{
                const province = document.getElementById('provinceInput').value.trim();
                if (!province) {{ alert('กรุณาเลือกจังหวัด'); return; }}
                
                const errors = document.querySelectorAll('.error-text');
                let hasCalcError = false;
                errors.forEach(err => {{ if(err.style.display === 'block') hasCalcError = true; }});
                if(hasCalcError) {{ alert('❌ พบข้อผิดพลาด! มียอดนักเรียนพิเศษรวมกันมากกว่ายอดนักเรียนทั้งหมด กรุณาตรวจสอบช่องตัวแดง'); return; }}

                const payload = {{ province: province, dlas: [] }};
                const dlaBlocks = document.querySelectorAll('.dla-block');
                let hasError = false;
                
                dlaBlocks.forEach(dlaBlock => {{
                    const dlaName = dlaBlock.querySelector('.dla-name').value.trim();
                    if (!dlaName) hasError = true;
                    const schools = [];
                    
                    dlaBlock.querySelectorAll('.school-wrapper').forEach(wrapper => {{
                        const schoolName = wrapper.querySelector('.school-name').value.trim();
                        const rtCount = parseInt(wrapper.querySelector('.rt-count').value);
                        const ntCount = parseInt(wrapper.querySelector('.nt-count').value);
                        if (!schoolName) hasError = true;
                        
                        const hasSp = wrapper.querySelector('.has-special').checked;
                        const rtSp = {{}}; const ntSp = {{}};
                        
                        if(hasSp) {{
                            ['t1','t2','t3','t4','t5','t6','t7','t8','t9'].forEach(k => {{
                                rtSp[k] = parseInt(wrapper.querySelector(`.rt-sp-${{k}}`).value) || 0;
                                ntSp[k] = parseInt(wrapper.querySelector(`.nt-sp-${{k}}`).value) || 0;
                            }});
                        }}

                        schools.push({{ 
                            school_name: schoolName, 
                            rt_count: isNaN(rtCount) ? 0 : rtCount, 
                            nt_count: isNaN(ntCount) ? 0 : ntCount,
                            has_special: hasSp,
                            rt_special: rtSp,
                            nt_special: ntSp
                        }});
                    }});
                    payload.dlas.push({{ dla_name: dlaName, schools: schools }});
                }});
                
                if (hasError) {{ alert('กรุณาเลือก อปท. และชื่อโรงเรียนให้ครบถ้วน'); return; }}
                if (payload.dlas.length === 0 || payload.dlas[0].schools.length === 0) {{ alert('ต้องมีข้อมูลอย่างน้อย 1 อปท. และ 1 โรงเรียน'); return; }}
                
                try {{
                    const response = await fetch('/submit-batch', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify(payload)
                    }});
                    const result = await response.json();
                    
                    if (result.status === "error") {{
                        alert(result.message);
                        window.location.href = "/dashboard?province=" + encodeURIComponent(province);
                    }} else if (result.status === "success") {{
                        document.body.innerHTML = `
                        <div class="container" style="text-align: center; margin-top: 50px;">
                            <h2 style="color: #38a169;">บันทึกข้อมูลเรียบร้อยแล้ว!</h2>
                            <p>ส่งข้อมูล ${{payload.dlas.length}} อปท. เข้าสู่ระบบสำเร็จ</p>
                            <a href="/" style="text-decoration: none; color: #2b6cb0; font-size: 16px; display: block; margin: 20px 0;">+ กลับไปหน้ากรอกข้อมูลใหม่</a>
                            <a href="/dashboard?province=${{encodeURIComponent(province)}}" style="text-decoration: none; color: #dd6b20; font-size: 18px; font-weight: bold;">📊 ตรวจสอบข้อมูลและปริ้นเอกสารรับรอง</a>
                        </div>`;
                    }}
                }} catch (error) {{ alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้'); }}
            }}
            
            window.onload = async () => {{ await fetchSchoolData(); }};
        </script>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)

@app.post("/submit-batch")
async def submit_batch(data: Submission):
    if is_province_locked(data.province):
        return {"status": "error", "message": "❌ จังหวัดนี้ยืนยันและล็อคข้อมูลแล้ว ไม่สามารถเพิ่มข้อมูลใหม่ได้!"}

    conn = get_db_connection()
    c = conn.cursor()
    for dla in data.dlas:
        for school in dla.schools:
            rt_json = json.dumps(school.rt_special) if school.has_special else '{}'
            nt_json = json.dumps(school.nt_special) if school.has_special else '{}'
            
            c.execute('''INSERT INTO school_data 
                         (province, dla_name, school_name, rt_student_count, nt_student_count, rt_special_json, nt_special_json) 
                         VALUES (%s, %s, %s, %s, %s, %s, %s)''',
                      (data.province, dla.dla_name, school.school_name, school.rt_count, school.nt_count, rt_json, nt_json))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/lock/{province}")
async def lock_province(province: str):
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT rt_student_count, nt_student_count FROM school_data WHERE province=%s", conn, params=(province,))
    if df.empty:
        conn.close()
        return RedirectResponse(url=f"/dashboard?province={province}", status_code=303)
        
    total_stu = int(df['rt_student_count'].sum() + df['nt_student_count'].sum())
    tz_th = timezone(timedelta(hours=7))
    now = datetime.now(tz_th)
    locked_at = now.strftime("%d/%m/%Y %H:%M:%S")
    ref_code = f"REF-{province}-{total_stu}-{now.strftime('%y%m%d%H%M')}"
    
    c = conn.cursor()
    c.execute("INSERT INTO province_locks (province, ref_code, locked_at) VALUES (%s, %s, %s) ON CONFLICT (province) DO UPDATE SET ref_code = EXCLUDED.ref_code, locked_at = EXCLUDED.locked_at", 
              (province, ref_code, locked_at))
    conn.commit()
    conn.close()
    
    return RedirectResponse(url=f"/dashboard?province={province}", status_code=303)

@app.get("/unlock/{province}")
async def unlock_province(province: str, key: str = ""):
    if key != "DLA2569":
        return HTMLResponse("<script>alert('รหัสผ่านไม่ถูกต้อง!'); window.history.back();</script>")
        
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT file_path FROM province_uploads WHERE province=%s", (province,))
    row = c.fetchone()
    if row and row[0]:
        try: supabase.storage.from_("signed-docs").remove([row[0]])
        except: pass
    c.execute("DELETE FROM province_uploads WHERE province=%s", (province,))
    c.execute("DELETE FROM province_locks WHERE province=%s", (province,))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/dashboard?province={province}", status_code=303)

@app.get("/dashboard", response_class=HTMLResponse)
async def view_dashboard(province: str = ""):
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT id, province, dla_name, school_name, rt_student_count, nt_student_count, rt_special_json, nt_special_json FROM school_data ORDER BY province, dla_name, school_name", conn)
    uploads = pd.read_sql_query("SELECT * FROM province_uploads", conn)
    upload_dict = uploads.set_index('province')['file_path'].to_dict()
    locks = pd.read_sql_query("SELECT * FROM province_locks", conn)
    lock_dict = locks.set_index('province').to_dict('index')
    conn.close()
    
    prov_opts = get_province_options(selected=province)
    prov_opts_clean = prov_opts.replace('<option value="">-- เลือกจังหวัด --</option>', '')
    
    if df.empty:
        table_html = "<p style='text-align: center; color: #666;'>ยังไม่มีข้อมูลการรายงานในระบบ</p>"
    else:
        table_html = '<table class="data-table" id="dataTable"><thead><tr><th>จังหวัด</th><th>อปท. สังกัด</th><th>โรงเรียน</th><th>ป.1 (RT)</th><th>ป.3 (NT)</th><th>จัดการ</th></tr></thead><tbody>'
        current_dla = None
        for index, row in df.iterrows():
            dla_display = row['dla_name'] if row['dla_name'] != current_dla else ""
            if row['dla_name'] != current_dla:
                if current_dla is not None:
                     table_html += '<tr class="dla-separator"><td colspan="6"></td></tr>'
                current_dla = row['dla_name']
                
            display_style = ""
            if province and row['province'] != province:
                display_style = "display: none;"
                
            is_locked = row['province'] in lock_dict
            action_btn = '<span style="color: #a0aec0; font-size: 13px;">🔒 ล็อคแล้ว</span>' if is_locked else f'<a href="/edit/{row["id"]}" class="btn-edit">✏️ แก้ไข/ลบ</a>'
            
            rt_sp_count = 0
            nt_sp_count = 0
            if pd.notna(row['rt_special_json']):
                val = row['rt_special_json']
                sp_data = json.loads(val) if isinstance(val, str) else val
                if isinstance(sp_data, dict): rt_sp_count = sum(sp_data.values())
            if pd.notna(row['nt_special_json']):
                val = row['nt_special_json']
                sp_data = json.loads(val) if isinstance(val, str) else val
                if isinstance(sp_data, dict): nt_sp_count = sum(sp_data.values())
                    
            has_sp = (rt_sp_count > 0) or (nt_sp_count > 0)
            sp_badge = f'<span style="color: #d69e2e; font-size: 12px; margin-left: 5px;" title="มีเด็กพิเศษ RT {rt_sp_count} คน / NT {nt_sp_count} คน">♿</span>' if has_sp else ''

            table_html += f'''
            <tr style="{display_style}">
                <td style="color: #718096; font-size: 12px;" class="prov-col">{row['province']}</td>
                <td style="font-weight: 500; color: #2c5282;">{dla_display}</td>
                <td>{row['school_name']}{sp_badge}</td>
                <td class="num">{row['rt_student_count']}</td>
                <td class="num">{row['nt_student_count']}</td>
                <td style="text-align: center;">{action_btn}</td>
            </tr>
            '''
        table_html += '</tbody></table>'
        
    action_html = ""
    if province:
        is_locked = province in lock_dict
        if not is_locked:
            action_html = f'''
            <div style="background: #fffaf0; border: 1px solid #feebc8; border-radius: 8px; padding: 20px; text-align: center; margin-bottom: 20px;">
                <h3 style="margin-top: 0; color: #dd6b20;">ส่วนที่ 2: การรับรองข้อมูลของจังหวัด {province}</h3>
                <p style="color: #c05621; font-size: 14px; margin-bottom: 15px;">⚠️ ข้อมูลยังไม่ถูกยืนยัน กดปุ่มด้านล่างเพื่อล็อคและพิมพ์เอกสาร</p>
                <form action="/lock/{province}" method="post">
                    <button type="submit" onclick="return confirm('ยืนยันระบบจะล็อคข้อมูลทันที?');" style="background-color: #dd6b20; color: white; padding: 10px 24px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: bold;">🔒 ยืนยันข้อมูลและสร้างรหัสอ้างอิง</button>
                </form>
            </div>
            '''
        else:
            ref_code = lock_dict[province]['ref_code']
            locked_at = lock_dict[province]['locked_at']
            action_html = f'''
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-bottom: 20px; position: relative;">
                <h3 style="margin-top: 0; color: #2b6cb0;">ส่วนที่ 2: การรับรองข้อมูลของจังหวัด {province}</h3>
                <div style="position: absolute; top: 20px; right: 20px; text-align: right; font-size: 12px; color: #718096; background: #edf2f7; padding: 5px 10px; border-radius: 4px;">
                    รหัส: <strong style="color: #2b6cb0;">{ref_code}</strong><br>เวลาล็อค: {locked_at}
                </div>
                <div style="display: flex; gap: 20px; margin-top: 30px;">
                    <div style="flex: 1; text-align: center;">
                        <p style="font-size: 14px;">1. ปริ้นเอกสาร</p>
                        <a href="/print/{province}" target="_blank" style="background-color: #d69e2e; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block;">🖨️ พิมพ์เอกสาร</a>
                    </div>
                    <div style="flex: 1; border-left: 1px solid #cbd5e0; padding-left: 20px;">
                        <p style="font-size: 14px;">2. อัปโหลดไฟล์สแกน</p>
            '''
            if province in upload_dict:
                try: file_url = supabase.storage.from_("signed-docs").get_public_url(upload_dict[province])
                except: file_url = "#"
                action_html += f'<div style="font-size: 14px; color: #38a169;">✅ <a href="{file_url}" target="_blank" style="color: #2b6cb0; font-weight: bold;">คลิกดูไฟล์รับรอง</a></div>'
            else:
                action_html += f'''
                        <form action="/upload/{province}" method="post" enctype="multipart/form-data" style="display: flex; gap: 10px;">
                            <input type="file" name="file" required accept=".pdf, .jpg, .png" style="font-size: 12px; width: 100%;">
                            <button type="submit" style="background-color: #3182ce; color: white; padding: 6px 12px; border: none; border-radius: 4px; cursor: pointer; white-space: nowrap;">📤 ยืนยันไฟล์</button>
                        </form>
                '''
            action_html += f'''
                    </div>
                </div>
            </div>
            '''

    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>ตรวจสอบข้อมูล - ระบบรายงาน RT/NT</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: #f0f2f5; padding: 20px; }}
            .container {{ max-width: 1000px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
            h2 {{ text-align: center; color: #2d3748; margin-bottom: 10px; }}
            .filter-section {{ display: flex; justify-content: space-between; margin-bottom: 15px; align-items: center; background: #ebf8ff; padding: 15px; border-radius: 8px; border: 1px solid #bee3f8; }}
            .filter-section select {{ padding: 8px 12px; border-radius: 4px; border: 1px solid #cbd5e0; font-family: 'Sarabun', sans-serif; width: 250px; }}
            .data-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }}
            .data-table th, .data-table td {{ border: 1px solid #e2e8f0; padding: 10px 12px; text-align: left; vertical-align: top; }}
            .data-table th {{ background-color: #2b6cb0; color: white; text-align: center; font-weight: 500; }}
            .data-table .num {{ text-align: center; font-weight: 500; }}
            .data-table .dla-separator td {{ background-color: #f8fafc; border: none; height: 10px; padding: 0; }}
            .nav-link {{ display: inline-block; margin-bottom: 15px; text-decoration: none; color: #2b6cb0; font-weight: 500; }}
            .btn-edit {{ background-color: #edf2f7; color: #2b6cb0; padding: 4px 10px; border-radius: 4px; text-decoration: none; font-size: 13px; border: 1px solid #cbd5e0; white-space: nowrap; display: inline-block; }}
            .btn-edit:hover {{ background-color: #e2e8f0; }}
        </style>
    </head>
    <body>
        <div class="container">
            <a href="/" class="nav-link">← กลับไปหน้าส่งข้อมูล (เพิ่มข้อมูล)</a>
            <h2>รายการข้อมูลที่จังหวัดบันทึกแล้ว</h2>
            <div class="filter-section">
                <div>
                    <h3 style="margin: 0; color: #2b6cb0; font-size: 16px;">ส่วนที่ 1: ตรวจสอบข้อมูล</h3>
                </div>
                <div>
                    <label style="font-weight: 500;">เลือกจังหวัด:</label>
                    <select id="provinceFilter" onchange="filterTable()">
                        <option value="">-- แสดงทุกจังหวัด --</option>
                        {prov_opts_clean}
                    </select>
                </div>
            </div>
            {action_html}
            {table_html}
            <div style="text-align: center; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 20px;">
                <a href="#" onclick="downloadExcel(event)" style="background-color: #38a169; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: 600;">📥 กรม: ดาวน์โหลด Excel ข้อมูลทั้งหมด</a>
            </div>
        </div>
        <script>
            function filterTable() {{
                var filter = document.getElementById("provinceFilter").value;
                if(filter) window.location.href = "/dashboard?province=" + encodeURIComponent(filter);
                else window.location.href = "/dashboard";
            }}
            function downloadExcel(e) {{
                e.preventDefault();
                let pass = prompt("รหัสผ่านดาวน์โหลด Excel:");
                if (pass) window.location.href = "/export?key=" + encodeURIComponent(pass);
            }}
        </script>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)

@app.get("/print/{province}", response_class=HTMLResponse)
async def print_page(province: str):
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT dla_name, school_name, rt_student_count, nt_student_count, rt_special_json, nt_special_json FROM school_data WHERE province=%s ORDER BY dla_name, school_name", conn, params=(province,))
    c = conn.cursor()
    c.execute("SELECT ref_code, locked_at FROM province_locks WHERE province=%s", (province,))
    lock_row = c.fetchone()
    conn.close()
    
    if df.empty: return HTMLResponse("<h2>ไม่พบข้อมูล</h2>")
    if not lock_row: return HTMLResponse("<h2>ข้อมูลยังไม่ถูกล็อค กรุณากดยืนยันข้อมูลก่อนพิมพ์เอกสาร</h2>")
    
    ref_code = lock_row[0]
    
    def ext_sp(j_data):
        if not j_data or j_data == '{}': return {}
        try: return json.loads(j_data) if isinstance(j_data, str) else j_data
        except: return {}

    table_rows = ""
    sum_rt = {'all':0, 'norm':0, 't1':0,'t2':0,'t3':0,'t4':0,'t5':0,'t6':0,'t7':0,'t8':0,'t9':0}
    sum_nt = {'all':0, 'norm':0, 't1':0,'t2':0,'t3':0,'t4':0,'t5':0,'t6':0,'t7':0,'t8':0,'t9':0}
    
    row_num = 1
    current_dla = None
    
    for index, row in df.iterrows():
        dla_display = row['dla_name'] if row['dla_name'] != current_dla else ""
        if row['dla_name'] != current_dla: current_dla = row['dla_name']
        
        rt_all = row['rt_student_count']
        nt_all = row['nt_student_count']
        
        rt_sp_d = ext_sp(row['rt_special_json'])
        nt_sp_d = ext_sp(row['nt_special_json'])
        
        rt_sp_tot = sum(rt_sp_d.values())
        nt_sp_tot = sum(nt_sp_d.values())
        
        rt_norm = rt_all - rt_sp_tot
        nt_norm = nt_all - nt_sp_tot
        
        sum_rt['all'] += rt_all; sum_rt['norm'] += rt_norm
        sum_nt['all'] += nt_all; sum_nt['norm'] += nt_norm
        
        for i in range(1, 10):
            sum_rt[f't{i}'] += rt_sp_d.get(f't{i}', 0)
            sum_nt[f't{i}'] += nt_sp_d.get(f't{i}', 0)

        table_rows += f'''
        <tr>
            <td class="center">{row_num}</td>
            <td class="truncate">{dla_display}</td>
            <td class="truncate">{row["school_name"]}</td>
            
            <td class="center" style="font-weight:bold; background:#f0f8ff;">{rt_all}</td>
            <td class="center">{rt_norm}</td>
            <td class="center sp-col">{rt_sp_d.get('t1', '')}</td><td class="center sp-col">{rt_sp_d.get('t2', '')}</td>
            <td class="center sp-col">{rt_sp_d.get('t3', '')}</td><td class="center sp-col">{rt_sp_d.get('t4', '')}</td>
            <td class="center sp-col">{rt_sp_d.get('t5', '')}</td><td class="center sp-col">{rt_sp_d.get('t6', '')}</td>
            <td class="center sp-col">{rt_sp_d.get('t7', '')}</td><td class="center sp-col">{rt_sp_d.get('t8', '')}</td>
            <td class="center sp-col">{rt_sp_d.get('t9', '')}</td>

            <td class="center" style="font-weight:bold; background:#f0faff;">{nt_all}</td>
            <td class="center">{nt_norm}</td>
            <td class="center sp-col">{nt_sp_d.get('t1', '')}</td><td class="center sp-col">{nt_sp_d.get('t2', '')}</td>
            <td class="center sp-col">{nt_sp_d.get('t3', '')}</td><td class="center sp-col">{nt_sp_d.get('t4', '')}</td>
            <td class="center sp-col">{nt_sp_d.get('t5', '')}</td><td class="center sp-col">{nt_sp_d.get('t6', '')}</td>
            <td class="center sp-col">{nt_sp_d.get('t7', '')}</td><td class="center sp-col">{nt_sp_d.get('t8', '')}</td>
            <td class="center sp-col">{nt_sp_d.get('t9', '')}</td>
        </tr>'''
        row_num += 1

    html_content = f'''<!DOCTYPE html>
    <html lang="th"><head><meta charset="utf-8"><title>เอกสารรับรองข้อมูล จังหวัด{province}</title>
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">
    <style>
        body {{ font-family:'Sarabun',sans-serif; font-size: 12px; color: #000; }}
        table {{ width:100%; border-collapse:collapse; margin-top: 10px; font-size: 10px; }}
        th, td {{ border:1px solid black; padding:4px 3px; }}
        th {{ background-color: #e2e8f0; text-align: center; font-weight: 600; line-height: 1.2; }}
        .center {{ text-align:center; }}
        .truncate {{ max-width: 120px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .sp-col {{ color: #b7791f; }}
        .ref-box {{ float: right; border: 1px dashed #666; padding: 5px 10px; font-size: 11px; color: #333; }}
        @media print {{
            @page {{ size: A4 landscape; margin: 10mm; }}
            body {{ -webkit-print-color-adjust: exact; }}
            .no-print {{ display: none; }}
        }}
    </style>
    </head><body>
    <div style="width:100%; margin:auto;">
        <div class="no-print" style="margin-bottom: 15px; text-align: center;">
            <button onclick="window.print()" style="background: #3182ce; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-family: Sarabun; font-size: 16px;">🖨️ กดพิมพ์เอกสาร (แนวนอน)</button>
            <p style="color: #c53030;">* ระบบจะตั้งค่าเป็นกระดาษ A4 แนวนอน (Landscape) อัตโนมัติ</p>
        </div>
        
        <div class="ref-box">รหัสอ้างอิง: {ref_code}</div>
        <div style="clear: both;"></div>
        
        <h3 style="text-align:center; margin-top: 0; font-size: 14px;">รายละเอียดจำนวนนักเรียนที่เข้ารับการประเมินคุณภาพผู้เรียน (RT/NT) ประจำปีการศึกษา 2569<br>จังหวัด{province}</h3>
        
        <table>
            <thead>
                <tr>
                    <th rowspan="2" style="width: 2%;">ที่</th>
                    <th rowspan="2" style="width: 12%;">อปท.</th>
                    <th rowspan="2" style="width: 14%;">โรงเรียน</th>
                    <th colspan="11" style="background:#bee3f8;">ป.1 (สอบ RT)</th>
                    <th colspan="11" style="background:#c6f6d5;">ป.3 (สอบ NT)</th>
                </tr>
                <tr>
                    <th style="width:3%;">รวม</th><th style="width:3%;">ปกติ</th>
                    <th style="width:2%;">เห็น</th><th style="width:2%;">ได้ยิน</th><th style="width:2%;">ปัญญา</th><th style="width:2%;">กาย</th><th style="width:2%;">LD</th><th style="width:2%;">พูด</th><th style="width:2%;">อารมณ์</th><th style="width:2%;">ออทิส</th><th style="width:2%;">ซ้อน</th>
                    <th style="width:3%;">รวม</th><th style="width:3%;">ปกติ</th>
                    <th style="width:2%;">เห็น</th><th style="width:2%;">ได้ยิน</th><th style="width:2%;">ปัญญา</th><th style="width:2%;">กาย</th><th style="width:2%;">LD</th><th style="width:2%;">พูด</th><th style="width:2%;">อารมณ์</th><th style="width:2%;">ออทิส</th><th style="width:2%;">ซ้อน</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
                <tr style="background-color: #edf2f7; font-weight: bold; font-size: 11px;">
                    <td colspan="3" style="text-align:right; padding-right: 10px;">รวมทั้งสิ้น</td>
                    <td class="center">{sum_rt['all']}</td><td class="center">{sum_rt['norm']}</td>
                    <td class="center sp-col">{sum_rt['t1']}</td><td class="center sp-col">{sum_rt['t2']}</td><td class="center sp-col">{sum_rt['t3']}</td><td class="center sp-col">{sum_rt['t4']}</td><td class="center sp-col">{sum_rt['t5']}</td><td class="center sp-col">{sum_rt['t6']}</td><td class="center sp-col">{sum_rt['t7']}</td><td class="center sp-col">{sum_rt['t8']}</td><td class="center sp-col">{sum_rt['t9']}</td>
                    <td class="center">{sum_nt['all']}</td><td class="center">{sum_nt['norm']}</td>
                    <td class="center sp-col">{sum_nt['t1']}</td><td class="center sp-col">{sum_nt['t2']}</td><td class="center sp-col">{sum_nt['t3']}</td><td class="center sp-col">{sum_nt['t4']}</td><td class="center sp-col">{sum_nt['t5']}</td><td class="center sp-col">{sum_nt['t6']}</td><td class="center sp-col">{sum_nt['t7']}</td><td class="center sp-col">{sum_nt['t8']}</td><td class="center sp-col">{sum_nt['t9']}</td>
                </tr>
            </tbody>
        </table>
        
        <div style="margin-top:30px; float: right; text-align:center; padding-right:50px; font-size: 12px;">
            <p style="margin-bottom: 25px;">ขอรับรองว่าข้อมูลดังกล่าวถูกต้องเป็นความจริงทุกประการ</p>
            <p>(ลงชื่อ)........................................................</p>
            <p>(........................................................)</p>
            <p>ตำแหน่ง........................................................</p>
            <p>วันที่........./................./.............</p>
        </div>
        <div style="clear: both;"></div>
    </div>
    <script>window.onload = function() {{ window.print(); }};</script>
    </body></html>'''
    return HTMLResponse(content=html_content)

@app.post("/upload/{province}")
async def upload_file(province: str, file: UploadFile = File(...)):
    file_ext = os.path.splitext(file.filename)[1]
    save_filename = f"signed_{datetime.now().strftime('%Y%m%d%H%M%S')}{file_ext}"
    file_bytes = await file.read()
    try: supabase.storage.from_("signed-docs").upload(save_filename, file_bytes, {"content-type": file.content_type})
    except: pass
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("INSERT INTO province_uploads (province, file_path) VALUES (%s, %s) ON CONFLICT (province) DO UPDATE SET file_path = EXCLUDED.file_path", (province, save_filename))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/dashboard?province={province}", status_code=303)

@app.get("/delete/{school_id}")
async def delete_data(school_id: int):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("DELETE FROM school_data WHERE id=%s", (school_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/dashboard", status_code=303)

@app.get("/export")
async def export_data(key: str = ""):
    if key != "DLA2569":
        return HTMLResponse("<script>alert('รหัสผ่านไม่ถูกต้อง! ระบบสงวนสิทธิ์เฉพาะเจ้าหน้าที่กรมเท่านั้น'); window.history.back();</script>")
        
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT province, dla_name, school_name, rt_student_count, nt_student_count, rt_special_json, nt_special_json FROM school_data ORDER BY province, dla_name, school_name", conn)
    conn.close()
    if df.empty: return HTMLResponse("<h2>ไม่มีข้อมูล</h2>")

    # 1. ข้อมูล Sheet 1 (งบประมาณเรียงยาว ไม่มีแถว Subtotal คั่นกลางให้รำคาญใจ)
    export_rows = []
    row_num = 1
    excel_row = 6 
    
    # 2. ข้อมูล Sheet 2 (สรุประดับจังหวัด)
    summary_rows = []
    sum_row_num = 1

    for prov in df['province'].unique():
        prov_df = df[df['province'] == prov]
        
        # --- คำนวณสรุปจังหวัด ---
        rt_c = prov_df['rt_student_count'].sum()
        nt_c = prov_df['nt_student_count'].sum()
        
        pao_rt = 10000 if rt_c > 0 else 0
        pao_nt = 10000 if nt_c > 0 else 0
        
        dla_rt = 0; dla_nt = 0
        for dla in prov_df['dla_name'].unique():
            dla_df = prov_df[prov_df['dla_name'] == dla]
            if dla_df['rt_student_count'].sum() > 0: dla_rt += 1000
            if dla_df['nt_student_count'].sum() > 0: dla_nt += 1000
            
            # --- สร้างแถวลง Sheet 1 ---
            export_rows.append({
                'A': row_num, 'B': f"  {dla}", 'C': None, 'D': None, 'E': None, 'F': None,
                'G': (1000 if dla_df['rt_student_count'].sum()>0 else 0) + (1000 if dla_df['nt_student_count'].sum()>0 else 0), 
                'H': None, 'I': f"=SUM(D{excel_row},F{excel_row},G{excel_row},H{excel_row})"
            })
            row_num += 1; excel_row += 1
            
            for _, row in dla_df.iterrows():
                export_rows.append({
                    'A': row_num, 'B': f"    - {row['school_name']}", 
                    'C': row['rt_student_count'], 'D': f"=IF(C{excel_row}>0, 250+(C{excel_row}*12), 0)",
                    'E': row['nt_student_count'], 'F': f"=IF(E{excel_row}>0, 250+(E{excel_row}*14), 0)",
                    'G': None, 'H': None, 'I': f"=SUM(D{excel_row},F{excel_row},G{excel_row},H{excel_row})"
                })
                row_num += 1; excel_row += 1
                
        sch_rt = sum(250 + (x * 12) for x in prov_df['rt_student_count'] if x > 0)
        sch_nt = sum(250 + (x * 14) for x in prov_df['nt_student_count'] if x > 0)
        
        summary_rows.append({
            'ลำดับ': sum_row_num, 'จังหวัด': prov,
            'งบ สถจ. (RT)': pao_rt, 'งบ สถจ. (NT)': pao_nt,
            'งบ อปท. (RT)': dla_rt, 'งบ อปท. (NT)': dla_nt,
            'งบโรงเรียน (RT)': sch_rt, 'งบโรงเรียน (NT)': sch_nt,
            'รวมงบ RT ทั้งสิ้น': pao_rt + dla_rt + sch_rt,
            'รวมงบ NT ทั้งสิ้น': pao_nt + dla_nt + sch_nt,
            'รวมงบประมาณ (บาท)': (pao_rt + dla_rt + sch_rt) + (pao_nt + dla_nt + sch_nt)
        })
        sum_row_num += 1

    export_rows.append({
        'A': 'รวมทั้งสิ้น', 'B': '', 
        'C': f"=SUM(C6:C{excel_row-1})", 'D': f"=SUM(D6:D{excel_row-1})",
        'E': f"=SUM(E6:E{excel_row-1})", 'F': f"=SUM(F6:F{excel_row-1})", 
        'G': f"=SUM(G6:G{excel_row-1})", 'H': f"=SUM(H6:H{excel_row-1})", 
        'I': f"=SUM(I6:I{excel_row-1})"
    })

    df_export = pd.DataFrame(export_rows)
    df_summary = pd.DataFrame(summary_rows)
    
    # 3. ข้อมูล Sheet 3 (รายงานเด็กพิเศษ)
    raw_rows = []
    def ext_sp(j_data):
        if not j_data or j_data == '{}': return {}
        try: return json.loads(j_data) if isinstance(j_data, str) else j_data
        except: return {}

    sp_row_num = 1
    for index, row in df.iterrows():
        rt_sp_d = ext_sp(row['rt_special_json'])
        nt_sp_d = ext_sp(row['nt_special_json'])
        rt_sp = sum(rt_sp_d.values())
        nt_sp = sum(nt_sp_d.values())
        
        raw_rows.append([
            sp_row_num, row['province'], row['dla_name'], row['school_name'],
            row['rt_student_count'], row['rt_student_count'] - rt_sp, rt_sp,
            rt_sp_d.get('t1', 0), rt_sp_d.get('t2', 0), rt_sp_d.get('t3', 0), rt_sp_d.get('t4', 0), rt_sp_d.get('t5', 0), rt_sp_d.get('t6', 0), rt_sp_d.get('t7', 0), rt_sp_d.get('t8', 0), rt_sp_d.get('t9', 0),
            row['nt_student_count'], row['nt_student_count'] - nt_sp, nt_sp,
            nt_sp_d.get('t1', 0), nt_sp_d.get('t2', 0), nt_sp_d.get('t3', 0), nt_sp_d.get('t4', 0), nt_sp_d.get('t5', 0), nt_sp_d.get('t6', 0), nt_sp_d.get('t7', 0), nt_sp_d.get('t8', 0), nt_sp_d.get('t9', 0)
        ])
        sp_row_num += 1
        
    df_raw = pd.DataFrame(raw_rows)
    file_path = "export_rt_nt_2569_calculated.xlsx"
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # ----- Sheet 1 -----
        df_export.to_excel(writer, index=False, header=False, startrow=5, sheet_name='งบประมาณ_ระดับโรงเรียน')
        ws1 = writer.sheets['งบประมาณ_ระดับโรงเรียน']
        
        ws1.merge_cells('A1:I1'); ws1['A1'] = 'รายละเอียดประกอบการจัดสรรงบประมาณรายจ่ายประจำปีงบประมาณ พ.ศ. 2569'
        ws1['A1'].font = Font(bold=True, size=12); ws1['A1'].alignment = Alignment(horizontal='center')
        ws1.merge_cells('A2:I2'); ws1['A2'] = 'แผนงานยุทธศาสตร์พัฒนาบริการประชาชนและการพัฒนาประสิทธิภาพภาครัฐ งบดำเนินงาน'
        ws1['A2'].font = Font(bold=True, size=12); ws1['A2'].alignment = Alignment(horizontal='center')
        ws1.merge_cells('A3:I3'); ws1['A3'] = 'โครงการประเมินคุณภาพนักเรียนระดับการศึกษาภาคบังคับ ปีการศึกษา 2569'
        ws1['A3'].font = Font(bold=True, size=12); ws1['A3'].alignment = Alignment(horizontal='center')

        ws1.merge_cells('A4:A5'); ws1['A4'] = 'ลำดับ'
        ws1.merge_cells('B4:B5'); ws1['B4'] = 'อปท./โรงเรียน'
        ws1.merge_cells('C4:D4'); ws1['C4'] = 'การสอบ RT (ชั้น ป.1)'; ws1['C5'] = 'นักเรียน (คน)'; ws1['D5'] = 'งบโรงเรียน\n(250 + 12/คน)'
        ws1.merge_cells('E4:F4'); ws1['E4'] = 'การสอบ NT (ชั้น ป.3)'; ws1['E5'] = 'นักเรียน (คน)'; ws1['F5'] = 'งบโรงเรียน\n(250 + 14/คน)'
        ws1.merge_cells('G4:G5'); ws1['G4'] = 'งบ อปท.\n(RT 1,000 / NT 1,000)'
        ws1.merge_cells('H4:H5'); ws1['H4'] = 'งบ สถจ.\n(RT 10,000 / NT 10,000)'
        ws1.merge_cells('I4:I5'); ws1['I4'] = 'รวมทั้งสิ้น\n(บาท)'
        
        ws1.column_dimensions['A'].width = 8; ws1.column_dimensions['B'].width = 40; ws1.column_dimensions['C'].width = 15; ws1.column_dimensions['D'].width = 25
        ws1.column_dimensions['E'].width = 15; ws1.column_dimensions['F'].width = 25; ws1.column_dimensions['G'].width = 25; ws1.column_dimensions['H'].width = 25; ws1.column_dimensions['I'].width = 20

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for r in range(4, 6):
            for c in range(1, 10):
                cell = ws1.cell(row=r, column=c)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                cell.font = Font(bold=True); cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = thin_border
        
        for r in range(6, len(export_rows) + 6):
            for c in range(1, 10):
                cell = ws1.cell(row=r, column=c)
                cell.border = thin_border
                if c == 1: cell.alignment = Alignment(horizontal='center', vertical='center')
                elif c == 2: cell.alignment = Alignment(horizontal='left', vertical='center')
                else: 
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    if cell.value is not None: cell.number_format = '#,##0'
            if ws1.cell(row=r, column=1).value == 'รวมทั้งสิ้น':
                for c in range(1, 10): ws1.cell(row=r, column=c).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"); ws1.cell(row=r, column=c).font = Font(bold=True)
                
        # ----- Sheet 2 (สรุประดับจังหวัด) -----
        df_summary.to_excel(writer, index=False, sheet_name='สรุประดับจังหวัด')
        ws2 = writer.sheets['สรุประดับจังหวัด']
        for col in ws2.columns:
            col_letter = col[0].column_letter
            ws2.column_dimensions[col_letter].width = 20
            for cell in col:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                    cell.font = Font(bold=True)
                elif cell.column > 2:
                    cell.number_format = '#,##0'
                    
        # ----- Sheet 3 (ข้อมูลเด็กพิเศษ) -----
        df_raw.to_excel(writer, index=False, header=False, startrow=3, sheet_name='รายงานเด็กพิเศษ')
        ws_sp = writer.sheets['รายงานเด็กพิเศษ']
        ws_sp.merge_cells('A1:AB1'); ws_sp['A1'] = 'รายงานสรุปจำนวนนักเรียนที่มีความต้องการจำเป็นพิเศษ (เรียนร่วม) ปีการศึกษา 2569'; ws_sp['A1'].font = Font(bold=True, size=14); ws_sp['A1'].alignment = Alignment(horizontal='center')
        ws_sp.merge_cells('A2:D2'); ws_sp['A2'] = 'ข้อมูลสถานศึกษา'
        ws_sp.merge_cells('E2:P2'); ws_sp['E2'] = 'ระดับชั้น ป.1 (สอบ RT)'
        ws_sp.merge_cells('Q2:AB2'); ws_sp['Q2'] = 'ระดับชั้น ป.3 (สอบ NT)'
        
        headers = ['ลำดับ', 'จังหวัด', 'อปท.', 'โรงเรียน', 'รวม', 'ปกติ', 'พิเศษรวม', 'เห็น', 'ได้ยิน', 'ปัญญา', 'ร่างกาย', 'LD', 'พูด/ภาษา', 'พฤติกรรม', 'ออทิสติก', 'ซ้อน', 'รวม', 'ปกติ', 'พิเศษรวม', 'เห็น', 'ได้ยิน', 'ปัญญา', 'ร่างกาย', 'LD', 'พูด/ภาษา', 'พฤติกรรม', 'ออทิสติก', 'ซ้อน']
        for col_num, header in enumerate(headers, 1): ws_sp.cell(row=3, column=col_num).value = header
            
        for col in range(1, 29):
            ws_sp.cell(row=2, column=col).border = thin_border; ws_sp.cell(row=3, column=col).border = thin_border
            ws_sp.cell(row=3, column=col).font = Font(bold=True); ws_sp.cell(row=2, column=col).font = Font(bold=True)
            ws_sp.cell(row=2, column=col).alignment = Alignment(horizontal='center', vertical='center'); ws_sp.cell(row=3, column=col).alignment = Alignment(horizontal='center', vertical='center')
            if col <= 4: ws_sp.cell(row=2, column=col).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"); ws_sp.cell(row=3, column=col).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            elif col <= 16: ws_sp.cell(row=2, column=col).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"); ws_sp.cell(row=3, column=col).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            else: ws_sp.cell(row=2, column=col).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"); ws_sp.cell(row=3, column=col).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
        ws_sp.column_dimensions['B'].width = 15; ws_sp.column_dimensions['C'].width = 25; ws_sp.column_dimensions['D'].width = 30
        for col_letter in ['E','F','G', 'Q','R','S']: ws_sp.column_dimensions[col_letter].width = 10
        
        for r in range(4, len(raw_rows) + 4):
            for c in range(1, 29):
                cell = ws_sp.cell(row=r, column=c)
                cell.border = thin_border
                if c > 4: cell.alignment = Alignment(horizontal='center')
                if c > 4 and cell.value == 0: cell.font = Font(color="CCCCCC")

    return FileResponse(file_path, filename="สรุปงบประมาณ_RT_NT_2569.xlsx")

@app.get("/edit/{school_id}", response_class=HTMLResponse)
async def edit_page(school_id: int):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM school_data WHERE id=%s", (school_id,))
    row = c.fetchone()
    conn.close()
    
    if not row: return HTMLResponse("<h2>ไม่พบข้อมูล</h2>")
    
    prov = row[1]
    if is_province_locked(prov): return HTMLResponse(f"<script>alert('จังหวัด {prov} ถูกล็อคแล้ว ไม่สามารถแก้ไขได้'); window.location.href='/dashboard?province={prov}';</script>")

    rt_sp_data = json.loads(row[6]) if isinstance(row[6], str) else (row[6] or {})
    nt_sp_data = json.loads(row[7]) if isinstance(row[7], str) else (row[7] or {})
    has_special = bool(rt_sp_data) or bool(nt_sp_data)
    checked_str = "checked" if has_special else ""
    panel_display = "block" if has_special else "none"

    def get_sp_val(data_dict, key): return data_dict.get(key, 0)
        
    rt_inputs = "".join([f'<div class="sp-row"><label>{name}</label><input type="number" name="rt_sp_{cid}" class="sp-input rt-sp-{cid}" min="0" value="{get_sp_val(rt_sp_data, cid)}" oninput="calcSpecial(\'rt\')"></div>' for cid, name in SPECIAL_CATEGORIES])
    nt_inputs = "".join([f'<div class="sp-row"><label>{name}</label><input type="number" name="nt_sp_{cid}" class="sp-input nt-sp-{cid}" min="0" value="{get_sp_val(nt_sp_data, cid)}" oninput="calcSpecial(\'nt\')"></div>' for cid, name in SPECIAL_CATEGORIES])

    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>แก้ไขข้อมูลโรงเรียน</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: #f0f2f5; padding: 20px; }}
            .container {{ max-width: 800px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
            h2 {{ text-align: center; color: #2b6cb0; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; margin-bottom: 20px; }}
            label {{ font-weight: 500; display: block; margin-bottom: 5px; font-size: 14px; }}
            input, select {{ width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #cbd5e0; border-radius: 6px; box-sizing: border-box; font-family: 'Sarabun'; }}
            .flex-row {{ display: flex; gap: 15px; }}
            .flex-row > div {{ flex: 1; }}
            .special-toggle {{ background: #ebf8ff; padding: 10px 15px; border-radius: 6px; display: inline-block; cursor: pointer; margin-bottom: 15px; width: 100%; box-sizing: border-box; font-weight: 500; }}
            .special-toggle input {{ width: auto; margin-right: 10px; transform: scale(1.2); }}
            .special-panel {{ display: {panel_display}; padding-top: 15px; border-top: 1px dashed #cbd5e0; }}
            .sp-grid {{ display: flex; gap: 20px; }}
            .sp-col {{ flex: 1; background: #fdfaf4; padding: 15px; border-radius: 8px; border: 1px solid #f6e05e; }}
            .sp-col-title {{ font-weight: 600; text-align: center; color: #b7791f; margin-bottom: 10px; border-bottom: 1px solid #fbd38d; padding-bottom: 5px; }}
            .sp-row {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; }}
            .sp-row label {{ font-size: 13px; font-weight: 400; margin-bottom: 0; flex: 2; }}
            .sp-row input {{ flex: 1; padding: 4px 8px; text-align: center; margin-bottom: 0; }}
            .sp-summary {{ margin-top: 10px; padding: 8px; background: white; border-radius: 4px; text-align: center; font-size: 14px; font-weight: 600; box-shadow: inset 0 1px 3px rgba(0,0,0,0.1); }}
            .error-text {{ color: #e53e3e; font-size: 13px; margin-top: 5px; display: none; font-weight: 500; text-align: center; }}
            .btn-group {{ display: flex; justify-content: space-between; margin-top: 30px; border-top: 1px solid #e2e8f0; padding-top: 20px; }}
            .btn {{ padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-family: 'Sarabun'; font-weight: 600; font-size: 15px; text-decoration: none; text-align: center; }}
            .btn-save {{ background-color: #3182ce; color: white; flex: 2; margin-right: 10px; }}
            .btn-del {{ background-color: #e53e3e; color: white; flex: 1; margin-right: 10px; }}
            .btn-back {{ background-color: #edf2f7; color: #4a5568; border: 1px solid #cbd5e0; flex: 1; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>✏️ แก้ไขข้อมูลโรงเรียน</h2>
            <div style="background: #f8fafc; padding: 15px; border-radius: 8px; margin-bottom: 20px; font-size: 14px; color: #4a5568;">
                <strong>จังหวัด:</strong> {row[1]} <br>
                <strong>อปท. ต้นสังกัด:</strong> {row[2]}
            </div>
            
            <form action="/update/{school_id}" method="post" onsubmit="return validateForm()">
                <label>ชื่อโรงเรียน</label>
                <input type="text" name="school_name" value="{row[3]}" required>
                
                <div class="flex-row">
                    <div>
                        <label>ยอด นร. ป.1 (RT) รวม</label>
                        <input type="number" id="rt_count" name="rt_count" min="0" value="{row[4]}" oninput="calcSpecial('rt')" required>
                    </div>
                    <div>
                        <label>ยอด นร. ป.3 (NT) รวม</label>
                        <input type="number" id="nt_count" name="nt_count" min="0" value="{row[5]}" oninput="calcSpecial('nt')" required>
                    </div>
                </div>
                
                <label class="special-toggle">
                    <input type="checkbox" name="has_special" id="has_special" {checked_str} onchange="toggleSpecialPanel(this)"> 
                    ♿ โรงเรียนนี้มีนักเรียนที่มีความต้องการจำเป็นพิเศษ (เรียนร่วม)
                </label>
                
                <div class="special-panel" id="special_panel">
                    <div class="sp-grid">
                        <div class="sp-col">
                            <div class="sp-col-title">ระบุประเภทพิเศษ ป.1 (RT)</div>
                            {rt_inputs}
                            <div class="error-text" id="rt_error">❌ ยอดเด็กพิเศษรวมกัน มากกว่ายอดทั้งหมด!</div>
                            <div class="sp-summary" id="rt_summary">จำนวนเด็กปกติ: <span id="rt_normal_num">0</span> คน</div>
                        </div>
                        <div class="sp-col">
                            <div class="sp-col-title">ระบุประเภทพิเศษ ป.3 (NT)</div>
                            {nt_inputs}
                            <div class="error-text" id="nt_error">❌ ยอดเด็กพิเศษรวมกัน มากกว่ายอดทั้งหมด!</div>
                            <div class="sp-summary" id="nt_summary">จำนวนเด็กปกติ: <span id="nt_normal_num">0</span> คน</div>
                        </div>
                    </div>
                </div>
                
                <div class="btn-group">
                    <button type="submit" class="btn btn-save">💾 บันทึกการแก้ไข</button>
                    <a href="/delete/{school_id}" class="btn btn-del" onclick="return confirm('ยืนยันการลบข้อมูลโรงเรียนนี้?');">🗑️ ลบทิ้ง</a>
                    <a href="/dashboard?province={row[1]}" class="btn btn-back">ยกเลิก</a>
                </div>
            </form>
        </div>
        <script>
            function toggleSpecialPanel(cb) {{
                const panel = document.getElementById('special_panel');
                panel.style.display = cb.checked ? 'block' : 'none';
                if(!cb.checked) {{
                    document.querySelectorAll('.sp-input').forEach(inp => inp.value = 0);
                    calcSpecial('rt'); calcSpecial('nt');
                }}
            }}
            function calcSpecial(type) {{
                const totalInput = document.getElementById(type + '_count');
                const total = parseInt(totalInput.value) || 0;
                let spTotal = 0;
                document.querySelectorAll('.' + type + '-sp-t1, .' + type + '-sp-t2, .' + type + '-sp-t3, .' + type + '-sp-t4, .' + type + '-sp-t5, .' + type + '-sp-t6, .' + type + '-sp-t7, .' + type + '-sp-t8, .' + type + '-sp-t9').forEach(inp => {{
                    spTotal += parseInt(inp.value) || 0;
                }});
                const normalCount = total - spTotal;
                const summaryBox = document.getElementById(type + '_summary');
                const errorBox = document.getElementById(type + '_error');
                if(normalCount < 0) {{
                    summaryBox.style.display = 'none'; errorBox.style.display = 'block'; totalInput.style.borderColor = '#e53e3e';
                }} else {{
                    summaryBox.style.display = 'block'; errorBox.style.display = 'none'; totalInput.style.borderColor = '#cbd5e0';
                    document.getElementById(type + '_normal_num').innerText = normalCount;
                }}
            }}
            function validateForm() {{
                if(document.getElementById('rt_error').style.display === 'block' || document.getElementById('nt_error').style.display === 'block') {{
                    alert('❌ มียอดเด็กพิเศษรวมกันมากกว่ายอดนักเรียนทั้งหมด กรุณาตรวจสอบตัวเลขอีกครั้ง'); return false;
                }}
                return true;
            }}
            window.onload = () => {{ calcSpecial('rt'); calcSpecial('nt'); }};
        </script>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)

@app.post("/update/{school_id}")
async def update_data(school_id: int, request: Request):
    form_data = await request.form()
    school_name = form_data.get("school_name")
    rt_count = int(form_data.get("rt_count", 0))
    nt_count = int(form_data.get("nt_count", 0))
    has_special = form_data.get("has_special") == "on"
    
    rt_sp = {}; nt_sp = {}
    if has_special:
        for cid, _ in SPECIAL_CATEGORIES:
            rt_val = int(form_data.get(f"rt_sp_{cid}", 0)); nt_val = int(form_data.get(f"nt_sp_{cid}", 0))
            if rt_val > 0: rt_sp[cid] = rt_val
            if nt_val > 0: nt_sp[cid] = nt_val
            
    conn = get_db_connection(); c = conn.cursor()
    c.execute("SELECT province FROM school_data WHERE id=%s", (school_id,))
    prov = c.fetchone()[0]
    c.execute('''UPDATE school_data SET school_name=%s, rt_student_count=%s, nt_student_count=%s, rt_special_json=%s, nt_special_json=%s WHERE id=%s''', 
              (school_name, rt_count, nt_count, json.dumps(rt_sp), json.dumps(nt_sp), school_id))
    conn.commit(); conn.close()
    return RedirectResponse(url=f"/dashboard?province={prov}", status_code=303)