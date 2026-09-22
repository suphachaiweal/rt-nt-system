import pandas as pd
import numpy as np
from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import List
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
# 🛑 กรุณาแก้ไขข้อมูล 2 บรรทัดนี้ให้เป็นของคุณ 🛑
# ==========================================
SUPABASE_URL = "https://svrtsuffmpoaeaozuihv.supabase.co" 
SUPABASE_KEY = "sb_publishable_S9c4kd6Ylxm2DUdzCOOOTQ_VVJV1GUY"
DB_URL = "postgresql://postgres:DlaAdmin2569Supabase@db.svrtsuffmpoaeaozuihv.supabase.co:5432/postgres"
# ==========================================

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

os.makedirs("uploads", exist_ok=True)
app.mount("/uploads", StaticFiles(directory="uploads"), name="uploads")

# ---- โหลดข้อมูลจากไฟล์ Excel ----
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
    print(f"Loaded {len(PROVINCES)} provinces from Excel.")
except Exception as e:
    print("Error loading Excel file:", e)
    DB_DATA = {}
    PROVINCES = []

@app.get("/api/school_data")
async def get_school_data():
    return JSONResponse(content=DB_DATA)
# ---------------------------------

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
                  
    c.execute('''CREATE TABLE IF NOT EXISTS province_uploads
                 (province TEXT PRIMARY KEY,
                  file_path TEXT)''')
                  
    c.execute('''CREATE TABLE IF NOT EXISTS province_locks
                 (province TEXT PRIMARY KEY,
                  ref_code TEXT,
                  locked_at TEXT)''')
    conn.commit()
    conn.close()

try:
    init_db()
    print("Connected to Supabase PostgreSQL successfully!")
except Exception as e:
    print("Database Connection Error:", e)

class School(BaseModel):
    school_name: str
    rt_count: int
    nt_count: int

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

@app.get("/", response_class=HTMLResponse)
async def get_form():
    prov_opts = get_province_options()
    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>ระบบส่งข้อมูลนักเรียนสอบ RT/NT ประจำปีการศึกษา 2569</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: #f0f2f5; padding: 20px; color: #1a202c; }}
            .container {{ max-width: 800px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }}
            h2 {{ text-align: center; color: #2d3748; margin-bottom: 25px; line-height: 1.4; }}
            .header-section {{ margin-bottom: 20px; }}
            .dla-block {{ background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-bottom: 20px; position: relative; }}
            .dla-title {{ font-weight: 600; font-size: 16px; margin-bottom: 15px; color: #2b6cb0; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; }}
            .school-block {{ display: flex; gap: 15px; margin-bottom: 10px; align-items: flex-end; background: white; padding: 10px; border-radius: 6px; border: 1px dashed #cbd5e0; }}
            .school-block > div {{ flex: 1; }}
            .school-block .col-school {{ flex: 2; }}
            label {{ font-weight: 500; font-size: 14px; margin-bottom: 5px; display: block; color: #4a5568; }}
            input, select {{ width: 100%; padding: 8px 12px; border: 1px solid #cbd5e0; border-radius: 4px; box-sizing: border-box; font-family: 'Sarabun', sans-serif; }}
            input:focus, select:focus {{ outline: none; border-color: #2b6cb0; box-shadow: 0 0 0 1px #2b6cb0; }}
            .btn {{ padding: 10px 15px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-family: 'Sarabun', sans-serif; font-weight: 500; transition: background 0.2s; }}
            .btn-primary {{ background-color: #2b6cb0; color: white; width: 100%; font-size: 16px; padding: 12px; margin-top: 20px; }}
            .btn-primary:hover {{ background-color: #2c5282; }}
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
                    <div class="school-block">
                        <div class="col-school">
                            <label>ชื่อโรงเรียน</label>
                            <select class="school-name" required><option value="">-- เลือกโรงเรียน --</option></select>
                        </div>
                        <div>
                            <label>นร. ป.1 (RT)</label>
                            <input type="number" class="rt-count" min="0" value="0" required>
                        </div>
                        <div>
                            <label>นร. ป.3 (NT)</label>
                            <input type="number" class="nt-count" min="0" value="0" required>
                        </div>
                        <div style="flex: 0.3; text-align: center; margin-bottom: 2px;">
                            <button type="button" class="btn btn-remove" onclick="this.parentElement.parentElement.remove()" title="ลบโรงเรียนนี้">ลบ</button>
                        </div>
                    </div>
                `;
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
                const payload = {{ province: province, dlas: [] }};
                const dlaBlocks = document.querySelectorAll('.dla-block');
                let hasError = false;
                dlaBlocks.forEach(dlaBlock => {{
                    const dlaName = dlaBlock.querySelector('.dla-name').value.trim();
                    if (!dlaName) hasError = true;
                    const schools = [];
                    dlaBlock.querySelectorAll('.school-block').forEach(schoolBlock => {{
                        const schoolName = schoolBlock.querySelector('.school-name').value.trim();
                        const rtCount = parseInt(schoolBlock.querySelector('.rt-count').value);
                        const ntCount = parseInt(schoolBlock.querySelector('.nt-count').value);
                        if (!schoolName) hasError = true;
                        schools.push({{ school_name: schoolName, rt_count: isNaN(rtCount) ? 0 : rtCount, nt_count: isNaN(ntCount) ? 0 : ntCount }});
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
            c.execute("INSERT INTO school_data (province, dla_name, school_name, rt_student_count, nt_student_count) VALUES (%s, %s, %s, %s, %s)",
                      (data.province, dla.dla_name, school.school_name, school.rt_count, school.nt_count))
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
        return HTMLResponse("<script>alert('รหัสผ่านไม่ถูกต้อง! ระบบสงวนสิทธิ์การปลดล็อคเฉพาะเจ้าหน้าที่กรมเท่านั้น'); window.history.back();</script>")
        
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT file_path FROM province_uploads WHERE province=%s", (province,))
    row = c.fetchone()
    
    if row and row[0]:
        try:
            supabase.storage.from_("signed-docs").remove([row[0]])
        except:
            pass
        
    c.execute("DELETE FROM province_uploads WHERE province=%s", (province,))
    c.execute("DELETE FROM province_locks WHERE province=%s", (province,))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/dashboard?province={province}", status_code=303)

@app.get("/dashboard", response_class=HTMLResponse)
async def view_dashboard(province: str = ""):
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT id, province, dla_name, school_name, rt_student_count, nt_student_count FROM school_data ORDER BY province, dla_name, school_name", conn)
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
        table_html = '<table class="data-table" id="dataTable"><thead><tr><th>จังหวัด</th><th>อปท. สังกัด</th><th>โรงเรียน</th><th>นร. ป.1 (RT)</th><th>นร. ป.3 (NT)</th><th>จัดการ</th></tr></thead><tbody>'
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
            if is_locked:
                action_btn = '<span style="color: #a0aec0; font-size: 13px;">🔒 ล็อคแล้ว</span>'
            else:
                action_btn = f'<a href="/edit/{row["id"]}" class="btn-edit">✏️ แก้ไข/ลบ</a>'
                
            table_html += f'''
            <tr style="{display_style}">
                <td style="color: #718096; font-size: 12px;" class="prov-col">{row['province']}</td>
                <td style="font-weight: 500; color: #2c5282;">{dla_display}</td>
                <td>{row['school_name']}</td>
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
            <div style="background: #fffaf0; border: 1px solid #feebc8; border-radius: 8px; padding: 20px; margin-top: 20px; text-align: center; margin-bottom: 20px;">
                <h3 style="margin-top: 0; color: #dd6b20;">ส่วนที่ 2: การรับรองข้อมูลของจังหวัด {province}</h3>
                <p style="color: #c05621; font-size: 14px; margin-bottom: 15px;">⚠️ ข้อมูลยังไม่ถูกยืนยัน คุณจะไม่สามารถพิมพ์เอกสารได้ จนกว่าจะกดปุ่มยืนยันด้านล่างนี้<br>(หากกดยืนยันแล้ว ระบบจะล็อคและไม่สามารถแก้ไขข้อมูลได้อีก)</p>
                <form action="/lock/{province}" method="post">
                    <button type="submit" onclick="return confirm('ตรวจสอบข้อมูลครบถ้วนแล้วใช่หรือไม่?\\nเมื่อยืนยันแล้วระบบจะล็อคข้อมูลทันที');" style="background-color: #dd6b20; color: white; padding: 10px 24px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: bold;">🔒 ยืนยันข้อมูลและสร้างรหัสอ้างอิง</button>
                </form>
            </div>
            '''
        else:
            ref_code = lock_dict[province]['ref_code']
            locked_at = lock_dict[province]['locked_at']
            
            action_html = f'''
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-top: 20px; margin-bottom: 20px; position: relative;">
                <h3 style="margin-top: 0; color: #2b6cb0;">ส่วนที่ 2: การรับรองข้อมูลของจังหวัด {province}</h3>
                <div style="position: absolute; top: 20px; right: 20px; text-align: right; font-size: 12px; color: #718096; background: #edf2f7; padding: 5px 10px; border-radius: 4px; border: 1px dashed #cbd5e0;">
                    รหัสอ้างอิง: <strong style="color: #2b6cb0;">{ref_code}</strong><br>เวลาล็อค: {locked_at}
                </div>
                <div style="display: flex; gap: 20px; align-items: flex-start; margin-top: 30px;">
                    <div style="flex: 1; text-align: center;">
                        <p style="margin-bottom: 10px; font-size: 14px; color: #4a5568;">1. ปริ้นเอกสารให้ผู้มีอำนาจลงนาม</p>
                        <a href="/print/{province}" target="_blank" style="background-color: #d69e2e; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block;">🖨️ พิมพ์เอกสารรับรองข้อมูล</a>
                    </div>
                    <div style="flex: 1; border-left: 1px solid #cbd5e0; padding-left: 20px;">
                        <p style="margin-bottom: 10px; font-size: 14px; color: #4a5568;">2. อัปโหลดไฟล์สแกนที่ลงนามแล้ว (PDF/JPG)</p>
            '''
            
            if province in upload_dict:
                try:
                    file_url = supabase.storage.from_("signed-docs").get_public_url(upload_dict[province])
                except:
                    file_url = "#"
                action_html += f'''
                        <div style="margin-top: 10px; font-size: 14px; color: #38a169;">
                            ✅ แนบไฟล์รับรองเรียบร้อยแล้ว: <a href="{file_url}" target="_blank" style="color: #2b6cb0; text-decoration: underline; font-weight: bold;">คลิกดูไฟล์รับรอง</a>
                        </div>
                '''
            else:
                action_html += f'''
                        <form action="/upload/{province}" method="post" enctype="multipart/form-data" style="display: flex; gap: 10px;">
                            <input type="file" name="file" required accept=".pdf, .jpg, .jpeg, .png" style="font-size: 12px; padding: 5px; width: 100%;">
                            <button type="submit" style="background-color: #3182ce; color: white; padding: 6px 12px; border: none; border-radius: 4px; cursor: pointer; white-space: nowrap;">📤 ยืนยันไฟล์</button>
                        </form>
                        <div style="margin-top: 10px; font-size: 13px; color: #e53e3e;">
                            ❌ ยังไม่ได้แนบไฟล์รับรอง
                        </div>
                '''
                
            action_html += f'''
                    </div>
                </div>
                <div style="margin-top: 25px; padding-top: 15px; border-top: 1px dashed #cbd5e0; text-align: center;">
                    <p style="font-size: 12px; color: #e53e3e; margin-bottom: 5px;">*ข้อมูลถูกล็อคแล้ว หากมีความจำเป็นต้องแก้ไขข้อมูล ให้แจ้งเจ้าหน้าที่กรมเพื่อปลดล็อคให้เท่านั้น</p>
                    <a href="#" onclick="unlockData('{province}')" style="background-color: #e53e3e; color: white; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-size: 13px; display: inline-block;">🔓 ปลดล็อคข้อมูล (เฉพาะเจ้าหน้าที่กรม)</a>
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
            .subtitle {{ text-align: center; color: #718096; margin-bottom: 25px; font-size: 14px; }}
            .filter-section {{ display: flex; justify-content: space-between; margin-bottom: 15px; align-items: center; background: #ebf8ff; padding: 15px; border-radius: 8px; border: 1px solid #bee3f8; }}
            .filter-section select {{ padding: 8px 12px; border-radius: 4px; border: 1px solid #cbd5e0; font-family: 'Sarabun', sans-serif; width: 250px; }}
            .data-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }}
            .data-table th, .data-table td {{ border: 1px solid #e2e8f0; padding: 10px 12px; text-align: left; vertical-align: top; }}
            .data-table th {{ background-color: #2b6cb0; color: white; text-align: center; font-weight: 500; }}
            .data-table .num {{ text-align: center; font-weight: 500; }}
            .data-table .dla-separator td {{ background-color: #f8fafc; border: none; height: 10px; padding: 0; }}
            .nav-link {{ display: inline-block; margin-bottom: 15px; text-decoration: none; color: #2b6cb0; font-weight: 500; }}
            .btn-edit {{ background-color: #edf2f7; color: #2b6cb0; padding: 4px 10px; border-radius: 4px; text-decoration: none; font-size: 13px; border: 1px solid #cbd5e0; }}
            .btn-edit:hover {{ background-color: #e2e8f0; }}
        </style>
    </head>
    <body>
        <div class="container">
            <a href="/" class="nav-link">← กลับไปหน้าส่งข้อมูล (เพิ่มข้อมูล)</a>
            <h2>รายการข้อมูลที่จังหวัดบันทึกแล้ว</h2>
            <div class="subtitle">ตรวจสอบความถูกต้อง ยืนยันข้อมูล และพิมพ์เอกสารรับรองสำหรับท้องถิ่นจังหวัด</div>
            
            <div class="filter-section">
                <div>
                    <h3 style="margin: 0 0 5px 0; color: #2b6cb0; font-size: 16px;">ส่วนที่ 1: ตรวจสอบข้อมูล</h3>
                    <div style="font-size: 13px; color: #4a5568;">กรุณาเลือกจังหวัดของท่านเพื่อดูข้อมูลและพิมพ์เอกสาร</div>
                </div>
                <div>
                    <label style="font-weight: 500; color: #4a5568; margin-right: 10px;">เลือกจังหวัด:</label>
                    <select id="provinceFilter" onchange="filterTable()">
                        <option value="">-- แสดงทุกจังหวัด --</option>
                        {prov_opts_clean}
                    </select>
                </div>
            </div>
            {action_html}
            {table_html}
            <div style="text-align: center; margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 20px;">
                <p style="color: #e53e3e; font-size: 13px; margin-bottom: 10px;">*สำหรับเจ้าหน้าที่ส่วนกลาง (กรม) เพื่อโหลดข้อมูลทั้งหมด</p>
                <a href="#" onclick="downloadExcel(event)" style="background-color: #38a169; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 16px; display: inline-block;">📥 กรม: ดาวน์โหลดไฟล์ Excel ข้อมูลทั้งหมด</a>
            </div>
        </div>
        <script>
            function filterTable() {{
                var filter = document.getElementById("provinceFilter").value;
                if(filter) {{ window.location.href = "/dashboard?province=" + encodeURIComponent(filter); }}
                else {{ window.location.href = "/dashboard"; }}
            }}
            function downloadExcel(e) {{
                e.preventDefault();
                let pass = prompt("ระบบสงวนสิทธิ์เฉพาะเจ้าหน้าที่กรม\\nกรุณากรอกรหัสผ่านเพื่อดาวน์โหลดไฟล์ Excel:");
                if (pass) {{ window.location.href = "/export?key=" + encodeURIComponent(pass); }}
            }}
            function unlockData(province) {{
                let pass = prompt("ระบบสงวนสิทธิ์เฉพาะเจ้าหน้าที่กรม\\nกรุณากรอกรหัสผ่านเพื่อปลดล็อคข้อมูล:");
                if (pass) {{ window.location.href = "/unlock/" + encodeURIComponent(province) + "?key=" + encodeURIComponent(pass); }}
            }}
        </script>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)

@app.get("/print/{province}", response_class=HTMLResponse)
async def print_page(province: str):
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT dla_name, school_name, rt_student_count, nt_student_count FROM school_data WHERE province=%s ORDER BY dla_name, school_name", conn, params=(province,))
    c = conn.cursor()
    c.execute("SELECT ref_code, locked_at FROM province_locks WHERE province=%s", (province,))
    lock_row = c.fetchone()
    conn.close()
    
    if df.empty:
        return HTMLResponse("<h2>ไม่พบข้อมูลของจังหวัดนี้</h2>")
    if not lock_row:
        return HTMLResponse("<h2 style='text-align: center; margin-top: 50px; color: red;'>ข้อมูลยังไม่ถูกยืนยัน (Lock)<br>กรุณากลับไปที่หน้าระบบและกดยืนยันข้อมูลก่อนพิมพ์เอกสาร</h2>")
        
    ref_code = lock_row[0]
    locked_at = lock_row[1]
    table_rows = ""
    total_rt = 0
    total_nt = 0
    row_num = 1
    current_dla = None
    
    for index, row in df.iterrows():
        dla_display = row['dla_name'] if row['dla_name'] != current_dla else ""
        if row['dla_name'] != current_dla:
             current_dla = row['dla_name']
        table_rows += f'''
        <tr>
            <td class="center">{row_num}</td>
            <td>{dla_display}</td>
            <td>{row['school_name']}</td>
            <td class="center">{row['rt_student_count']}</td>
            <td class="center">{row['nt_student_count']}</td>
        </tr>'''
        row_num += 1
        total_rt += row['rt_student_count']
        total_nt += row['nt_student_count']

    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>เอกสารรับรองข้อมูล จังหวัด{province}</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: white; color: black; font-size: 14px; padding: 20px; }}
            .page-container {{ max-width: 800px; margin: auto; position: relative; }}
            h2, h3 {{ text-align: center; margin-bottom: 5px; }}
            h3 {{ margin-bottom: 20px; font-weight: 500; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 30px; margin-bottom: 30px; }}
            th, td {{ border: 1px solid black; padding: 8px; vertical-align: top; }}
            th {{ background-color: #f0f0f0; text-align: center; }}
            .center {{ text-align: center; }}
            .right {{ text-align: right; }}
            .bold {{ font-weight: bold; }}
            .signature-box {{ margin-top: 60px; display: flex; justify-content: flex-end; }}
            .signature-content {{ text-align: center; width: 450px; }}
            @media print {{ .no-print {{ display: none; }} body {{ padding: 0; }} }}
        </style>
    </head>
    <body>
        <div class="no-print" style="text-align: center; margin-bottom: 30px;">
            <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #2b6cb0; color: white; border: none; border-radius: 4px;">🖨️ สั่งพิมพ์หน้านี้</button>
            <p style="color: #e53e3e; font-size: 13px;">(แนะนำให้ตั้งค่าการปริ้นเป็นกระดาษ A4 แนวตั้ง)</p>
        </div>
        <div class="page-container">
            <div style="text-align: right; margin-bottom: 15px;">
                <div style="display: inline-block; text-align: left; font-size: 12px; color: #555; border: 1px dashed #ccc; padding: 5px 10px;">
                    รหัสอ้างอิงเอกสาร: <strong style="color: #2b6cb0;">{ref_code}</strong><br>พิมพ์เมื่อ: {locked_at}
                </div>
            </div>
            <h2>รายละเอียดจำนวนนักเรียนที่เข้ารับการประเมินคุณภาพผู้เรียน (RT/NT)</h2>
            <h3>ประจำปีการศึกษา 2568<br>จังหวัด{province}</h3>
            <table>
                <thead>
                    <tr>
                        <th width="8%">ลำดับ</th><th width="25%">องค์กรปกครองส่วนท้องถิ่น</th><th width="35%">ชื่อโรงเรียน</th>
                        <th width="16%">จำนวน นร.<br>ป.1 (RT)</th><th width="16%">จำนวน นร.<br>ป.3 (NT)</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                    <tr>
                        <td colspan="3" class="right bold">รวมทั้งสิ้น</td>
                        <td class="center bold">{total_rt}</td><td class="center bold">{total_nt}</td>
                    </tr>
                </tbody>
            </table>
            <div class="signature-box">
                <div class="signature-content">
                    <p style="margin-bottom: 35px;">ขอรับรองว่าข้อมูลดังกล่าวถูกต้องเป็นความจริงทุกประการ</p>
                    <p>(ลงชื่อ)..........................................................................</p>
                    <p style="margin-top: 15px;">(..........................................................................)</p>
                    <p style="margin-top: 15px;">ตำแหน่ง..........................................................................</p>
                    <p style="margin-top: 15px;">ท้องถิ่นจังหวัด{province}</p>
                    <p style="margin-top: 15px;">วันที่........./................./.............</p>
                </div>
            </div>
        </div>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)

@app.post("/upload/{province}")
async def upload_file(province: str, file: UploadFile = File(...)):
    file_ext = os.path.splitext(file.filename)[1]
    save_filename = f"signed_{province}_{datetime.now().strftime('%Y%m%d%H%M%S')}{file_ext}"
    
    file_bytes = await file.read()
    try:
        res = supabase.storage.from_("signed-docs").upload(save_filename, file_bytes, {"content-type": file.content_type})
    except Exception as e:
        print("Upload Error:", e)
        with open(os.path.join("uploads", save_filename), "wb") as f:
            f.write(file_bytes)
        
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("INSERT INTO province_uploads (province, file_path) VALUES (%s, %s) ON CONFLICT (province) DO UPDATE SET file_path = EXCLUDED.file_path", 
              (province, save_filename))
    conn.commit()
    conn.close()
    
    return RedirectResponse(url=f"/dashboard?province={province}", status_code=303)

@app.get("/edit/{school_id}", response_class=HTMLResponse)
async def edit_page(school_id: int):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM school_data WHERE id=%s", (school_id,))
    row = c.fetchone()
    conn.close()
    
    if not row:
        return HTMLResponse("<h2>ไม่พบข้อมูลโรงเรียนนี้</h2><a href='/dashboard'>กลับไปหน้าตรวจสอบ</a>")

    if is_province_locked(row[1]):
        return HTMLResponse(f"<script>alert('จังหวัด {row[1]} ล็อคข้อมูลแล้ว ไม่สามารถแก้ไขได้ โปรดแจ้งส่วนกลางเพื่อปลดล็อค'); window.location.href='/dashboard?province={row[1]}';</script>")

    prov_opts = get_province_options(selected=row[1])

    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>แก้ไขข้อมูล - ระบบรายงาน RT/NT</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: #f4f7f6; padding: 20px; }}
            .container {{ max-width: 500px; margin: auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
            h2 {{ text-align: center; color: #333; }}
            label {{ font-weight: 500; margin-top: 10px; display: block; color: #4a5568; }}
            input, select {{ width: 100%; padding: 10px; margin: 8px 0 20px; border: 1px solid #cbd5e0; border-radius: 4px; box-sizing: border-box; font-family: 'Sarabun', sans-serif; }}
            .btn-group {{ display: flex; gap: 10px; margin-top: 10px; }}
            button {{ flex: 1; padding: 12px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; font-family: 'Sarabun', sans-serif; color: white; }}
            .btn-save {{ background-color: #38a169; }}
            .btn-save:hover {{ background-color: #2f855a; }}
            .btn-delete {{ background-color: #e53e3e; }}
            .btn-delete:hover {{ background-color: #c53030; }}
            .btn-cancel {{ background-color: #a0aec0; display: block; text-align: center; text-decoration: none; margin-top: 10px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>✏️ แก้ไขข้อมูลโรงเรียน</h2>
            <form action="/update/{row[0]}" method="post">
                <label>จังหวัด</label>
                <select name="province" id="editProvince" required onchange="updateEditDLA()">
                    {prov_opts}
                </select>
                <label>ชื่อ อปท.</label>
                <select name="dla_name" id="editDla" required onchange="updateEditSchool()">
                    <option value="{row[2]}">{row[2]}</option>
                </select>
                <label>ชื่อโรงเรียน</label>
                <select name="school_name" id="editSchool" required>
                    <option value="{row[3]}">{row[3]}</option>
                </select>
                <label>นร. ป.1 (RT)</label>
                <input type="number" name="rt_student_count" min="0" value="{row[4]}" required>
                <label>นร. ป.3 (NT)</label>
                <input type="number" name="nt_student_count" min="0" value="{row[5]}" required>
                
                <div class="btn-group">
                    <button type="submit" class="btn-save">💾 บันทึกการแก้ไข</button>
                    <button type="button" class="btn-delete" onclick="confirmDelete({row[0]})">🗑️ ลบโรงเรียนนี้</button>
                </div>
                <a href="/dashboard" class="btn-cancel" style="padding: 12px; border-radius: 4px; color: white;">❌ ยกเลิก</a>
            </form>
        </div>
        <script>
            let dbData = {{}};
            const currentDla = "{row[2]}";
            const currentSchool = "{row[3]}";
            
            async function fetchSchoolData() {{
                const res = await fetch('/api/school_data');
                dbData = await res.json();
                updateEditDLA();
            }}
            
            function updateEditDLA() {{
                const prov = document.getElementById('editProvince').value;
                const dlaSelect = document.getElementById('editDla');
                dlaSelect.innerHTML = '<option value="">-- เลือก อปท. --</option>';
                if (prov && dbData[prov]) {{
                    for (let d in dbData[prov]) {{
                        let sel = (d === currentDla) ? "selected" : "";
                        dlaSelect.innerHTML += `<option value="${{d}}" ${{sel}}>${{d}}</option>`;
                    }}
                }}
                updateEditSchool();
            }}
            
            function updateEditSchool() {{
                const prov = document.getElementById('editProvince').value;
                const dla = document.getElementById('editDla').value;
                const schoolSelect = document.getElementById('editSchool');
                schoolSelect.innerHTML = '<option value="">-- เลือกโรงเรียน --</option>';
                if (prov && dla && dbData[prov] && dbData[prov][dla]) {{
                    dbData[prov][dla].forEach(sch => {{
                        let sel = (sch === currentSchool) ? "selected" : "";
                        schoolSelect.innerHTML += `<option value="${{sch}}" ${{sel}}>${{sch}}</option>`;
                    }});
                }}
            }}
            
            window.onload = function() {{ fetchSchoolData(); }};
            function confirmDelete(id) {{
                if(confirm('ระวัง! คุณต้องการลบข้อมูลโรงเรียนนี้ใช่หรือไม่? (ลบแล้วกู้คืนไม่ได้)')) {{
                    window.location.href = '/delete/' + id;
                }}
            }}
        </script>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)

@app.post("/update/{school_id}")
async def update_data(
    school_id: int, province: str = Form(...), dla_name: str = Form(...),
    school_name: str = Form(...), rt_student_count: int = Form(...), nt_student_count: int = Form(...)
):
    if is_province_locked(province):
        return HTMLResponse(f"<script>alert('จังหวัด {province} ถูกล็อคแล้ว ไม่สามารถแก้ไขได้'); window.history.back();</script>")

    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''UPDATE school_data SET province=%s, dla_name=%s, school_name=%s, rt_student_count=%s, nt_student_count=%s WHERE id=%s''', 
              (province, dla_name, school_name, rt_student_count, nt_student_count, school_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/dashboard", status_code=303)

@app.get("/delete/{school_id}")
async def delete_data(school_id: int):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT province FROM school_data WHERE id=%s", (school_id,))
    prov = c.fetchone()
    if prov and is_province_locked(prov[0]):
        conn.close()
        return HTMLResponse(f"<script>alert('ไม่สามารถลบได้ จังหวัด {prov[0]} ถูกล็อคแล้ว'); window.history.back();</script>")

    c.execute("DELETE FROM school_data WHERE id=%s", (school_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/dashboard", status_code=303)

@app.get("/export")
async def export_data(key: str = ""):
    if key != "DLA2569":
        return HTMLResponse("<script>alert('รหัสผ่านไม่ถูกต้อง! ระบบสงวนสิทธิ์เฉพาะเจ้าหน้าที่กรมเท่านั้น'); window.history.back();</script>")
        
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT province, dla_name, school_name, rt_student_count, nt_student_count FROM school_data ORDER BY province, dla_name, school_name", conn)
    
    if df.empty:
        conn.close()
        return HTMLResponse("<h2>ยังไม่มีข้อมูลให้ดาวน์โหลด</h2><a href='/dashboard'>กลับไปหน้าตรวจสอบ</a>")

    prov_status = df.groupby('province').agg({'rt_student_count': 'sum', 'nt_student_count': 'sum'}).to_dict('index')
    df['prov_dla'] = df['province'] + "|" + df['dla_name']
    dla_status = df.groupby('prov_dla').agg({'rt_student_count': 'sum', 'nt_student_count': 'sum'}).to_dict('index')
    conn.close()

    export_rows = []
    current_province = None
    current_dla = None
    row_num = 1
    excel_row = 6 

    for index, row in df.iterrows():
        prov = row['province']
        dla = row['dla_name']
        sch = row['school_name']
        rt_c = row['rt_student_count']
        nt_c = row['nt_student_count']
        prov_dla_key = f"{prov}|{dla}"
        
        if prov != current_province:
            current_province = prov
            current_dla = None
            pao_rt_budget = 10000 if prov_status[prov]['rt_student_count'] > 0 else 0
            pao_nt_budget = 10000 if prov_status[prov]['nt_student_count'] > 0 else 0
            total_pao_budget = pao_rt_budget + pao_nt_budget
            
            export_rows.append({
                'A': row_num, 'B': f"{prov}", 'C': None, 'D': None, 'E': None, 'F': None,
                'G': None, 'H': total_pao_budget, 'I': f"=SUM(D{excel_row},F{excel_row},G{excel_row},H{excel_row})"
            })
            row_num += 1
            excel_row += 1
            
        if dla != current_dla:
            current_dla = dla
            dla_rt_budget = 1000 if dla_status[prov_dla_key]['rt_student_count'] > 0 else 0
            dla_nt_budget = 1000 if dla_status[prov_dla_key]['nt_student_count'] > 0 else 0
            total_dla_budget = dla_rt_budget + dla_nt_budget
            
            export_rows.append({
                'A': row_num, 'B': f"  {dla}", 'C': None, 'D': None, 'E': None, 'F': None,
                'G': total_dla_budget, 'H': None, 'I': f"=SUM(D{excel_row},F{excel_row},G{excel_row},H{excel_row})"
            })
            row_num += 1
            excel_row += 1
            
        export_rows.append({
            'A': row_num, 'B': f"    - {sch}", 'C': rt_c, 'D': f"=IF(C{excel_row}>0, 250+(C{excel_row}*12), 0)",
            'E': nt_c, 'F': f"=IF(E{excel_row}>0, 250+(E{excel_row}*14), 0)",
            'G': None, 'H': None, 'I': f"=SUM(D{excel_row},F{excel_row},G{excel_row},H{excel_row})"
        })
        row_num += 1
        excel_row += 1

    export_rows.append({
        'A': 'รวมทั้งสิ้น', 'B': '', 'C': f"=SUM(C6:C{excel_row-1})", 'D': f"=SUM(D6:D{excel_row-1})",
        'E': f"=SUM(E6:E{excel_row-1})", 'F': f"=SUM(F6:F{excel_row-1})", 'G': f"=SUM(G6:G{excel_row-1})",
        'H': f"=SUM(H6:H{excel_row-1})", 'I': f"=SUM(I6:I{excel_row-1})"
    })

    df_export = pd.DataFrame(export_rows)
    file_path = "export_rt_nt_2568_calculated.xlsx"
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, header=False, startrow=5, sheet_name='งบประมาณ_RT_NT')
        worksheet = writer.sheets['งบประมาณ_RT_NT']
        
        worksheet.merge_cells('A1:I1')
        worksheet['A1'] = 'รายละเอียดประกอบการจัดสรรงบประมาณรายจ่ายประจำปีงบประมาณ พ.ศ. 2569'
        worksheet['A1'].font = Font(bold=True, size=12)
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        
        worksheet.merge_cells('A2:I2')
        worksheet['A2'] = 'แผนงานยุทธศาสตร์พัฒนาบริการประชาชนและการพัฒนาประสิทธิภาพภาครัฐ งบดำเนินงาน'
        worksheet['A2'].font = Font(bold=True, size=12)
        worksheet['A2'].alignment = Alignment(horizontal='center', vertical='center')
        
        worksheet.merge_cells('A3:I3')
        worksheet['A3'] = 'โครงการประเมินคุณภาพนักเรียนระดับการศึกษาภาคบังคับ ปีการศึกษา 2568'
        worksheet['A3'].font = Font(bold=True, size=12)
        worksheet['A3'].alignment = Alignment(horizontal='center', vertical='center')

        worksheet.merge_cells('A4:A5')
        worksheet['A4'] = 'ลำดับ'
        
        worksheet.merge_cells('B4:B5')
        worksheet['B4'] = 'จังหวัด/อปท./โรงเรียน'
        
        worksheet.merge_cells('G4:G5')
        worksheet['G4'] = 'งบ อปท.\n(RT 1,000 / NT 1,000)'
        
        worksheet.merge_cells('H4:H5')
        worksheet['H4'] = 'งบ สถจ.\n(RT 10,000 / NT 10,000)'
        
        worksheet.merge_cells('I4:I5')
        worksheet['I4'] = 'รวมทั้งสิ้น\n(บาท)'
        
        worksheet.merge_cells('C4:D4')
        worksheet['C4'] = 'การสอบ RT (ชั้น ป.1)'
        worksheet['C5'] = 'จำนวนนักเรียน\n(คน)'
        worksheet['D5'] = 'งบโรงเรียน\n(250บ. + 12บ./คน)'
        
        worksheet.merge_cells('E4:F4')
        worksheet['E4'] = 'การสอบ NT (ชั้น ป.3)'
        worksheet['E5'] = 'จำนวนนักเรียน\n(คน)'
        worksheet['F5'] = 'งบโรงเรียน\n(250บ. + 14บ./คน)'
        
        worksheet.column_dimensions['A'].width = 8
        worksheet.column_dimensions['B'].width = 40
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 25
        worksheet.column_dimensions['E'].width = 15
        worksheet.column_dimensions['F'].width = 25
        worksheet.column_dimensions['G'].width = 25
        worksheet.column_dimensions['H'].width = 25
        worksheet.column_dimensions['I'].width = 20

        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for row in range(4, 6):
            for col in range(1, 10):
                cell = worksheet.cell(row=row, column=col)
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = thin_border
        
        last_row = len(export_rows) + 5
        for row in range(6, last_row + 1):
            for col in range(1, 10):
                cell = worksheet.cell(row=row, column=col)
                cell.border = thin_border
                if col == 1:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col == 2:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0'

        for col in range(1, 10):
            cell = worksheet.cell(row=last_row, column=col)
            cell.font = Font(bold=True)
            cell.fill = total_fill

    return FileResponse(file_path, filename="สรุปงบประมาณ_RT_NT_2568.xlsx")