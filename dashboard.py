from fastapi import Request
from fastapi.responses import HTMLResponse
import pandas as pd
import json
from main import app, get_db_connection, PROVINCES  # ดึงระบบเดิมจาก main.py มาทำงานร่วมกัน

@app.get("/overview", response_class=HTMLResponse)
async def view_public_dashboard():
    # 1. เชื่อมต่อฐานข้อมูล
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT province, dla_name, school_name, rt_student_count, nt_student_count, rt_special_json, nt_special_json FROM school_data", conn)
    locks = pd.read_sql_query("SELECT province FROM province_locks", conn)
    uploads = pd.read_sql_query("SELECT province FROM province_uploads", conn)
    conn.close()
    
    # 2. คำนวณตัวเลขสถิติภาพรวม
    total_provinces = 76
    locked_count = len(locks)
    uploaded_count = len(uploads)
    
    locked_pct = round((locked_count / total_provinces) * 100, 2) if total_provinces else 0
    uploaded_pct = round((uploaded_count / total_provinces) * 100, 2) if total_provinces else 0
    
    total_schools = len(df)
    sent_provinces_count = df['province'].nunique() if not df.empty else 0
    
    rt_total, rt_special = 0, 0
    nt_total, nt_special = 0, 0
    
    if not df.empty:
        rt_total = int(df['rt_student_count'].sum())
        nt_total = int(df['nt_student_count'].sum())
        
        for _, row in df.iterrows():
            rt_sp_data = row['rt_special_json']
            if pd.notna(rt_sp_data):
                sp_dict = json.loads(rt_sp_data) if isinstance(rt_sp_data, str) else rt_sp_data
                if isinstance(sp_dict, dict): rt_special += sum(sp_dict.values())
                    
            nt_sp_data = row['nt_special_json']
            if pd.notna(nt_sp_data):
                sp_dict = json.loads(nt_sp_data) if isinstance(nt_sp_data, str) else nt_sp_data
                if isinstance(sp_dict, dict): nt_special += sum(sp_dict.values())
    
    rt_normal = rt_total - rt_special
    nt_normal = nt_total - nt_special

    # 3. วาดหน้าจอ HTML + CSS สไตล์แดชบอร์ด
    html_content = f'''<!DOCTYPE html>
    <html lang="th">
    <head>
        <meta charset="utf-8">
        <title>ภาพรวมการรายงานข้อมูล RT/NT</title>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Sarabun', sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; color: #333; }}
            .container {{ max-width: 1200px; margin: auto; }}
            .header-banner {{ text-align: center; margin-bottom: 20px; }}
            .header-banner h1 {{ font-size: 24px; color: #000; margin-bottom: 10px; }}
            .welcome-bar {{ background-color: #fdf5e6; border: 1px solid #fbeed5; color: #b7791f; text-align: center; padding: 10px; font-weight: bold; border-radius: 4px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
            
            .section-card {{ background: #fff; border: 1px solid #dee2e6; border-radius: 4px; margin-bottom: 20px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
            .section-header {{ background-color: #e9ecef; padding: 12px 20px; font-weight: bold; font-size: 15px; border-bottom: 1px solid #dee2e6; color: #495057; display: flex; align-items: center; gap: 10px; }}
            
            .table-layout {{ width: 100%; border-collapse: collapse; }}
            .table-layout td {{ padding: 12px 20px; border-bottom: 1px solid #eee; font-size: 14px; vertical-align: middle; }}
            .table-layout tr:last-child td {{ border-bottom: none; }}
            .td-label {{ width: 35%; color: #333; }}
            .td-value {{ width: 15%; text-align: center; font-weight: bold; }}
            .td-progress {{ width: 40%; }}
            .td-detail {{ width: 10%; text-align: right; }}
            .td-detail a {{ color: #dc3545; text-decoration: none; font-size: 12px; }}
            .td-detail a:hover {{ text-decoration: underline; }}
            
            .progress-bg {{ background-color: #e9ecef; border-radius: 4px; height: 12px; width: 100%; overflow: hidden; margin-top: 5px; }}
            .progress-bar-cyan {{ background-color: #17a2b8; height: 100%; }}
            .progress-bar-green {{ background-color: #28a745; height: 100%; }}
            
            .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
            
            .card-rt .section-header {{ background-color: #f8d7da; color: #721c24; border-bottom: 1px solid #f5c6cb; }}
            .card-rt .sub-header {{ background-color: #fdf3f4; font-weight: bold; padding: 10px 20px; border-bottom: 1px solid #eee; color: #721c24; }}
            
            .card-nt .section-header {{ background-color: #d4edda; color: #155724; border-bottom: 1px solid #c3e6cb; }}
            .card-nt .sub-header {{ background-color: #f3fcf5; font-weight: bold; padding: 10px 20px; border-bottom: 1px solid #eee; color: #155724; }}
            
            .data-row {{ display: flex; justify-content: space-between; padding: 10px 20px; border-bottom: 1px solid #eee; font-size: 14px; }}
            .data-row:nth-child(even) {{ background-color: #fcfcfc; }}
            .val-num {{ font-weight: 500; text-align: right; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-banner">
                <h1># ระบบรายงานและติดตามข้อมูลนักเรียน ป.1 และ ป.3 #</h1>
            </div>
            
            <div class="welcome-bar">
                “ ยินดีต้อนรับ องค์กรปกครองส่วนท้องถิ่น เข้าสู่ระบบ ”
            </div>
            
            <!-- Section 1: Progress -->
            <div class="section-card">
                <div class="section-header">📢 ข้อความแจ้งเตือนจากระบบ (ความคืบหน้าระดับประเทศ)</div>
                <table class="table-layout">
                    <tr>
                        <td class="td-label">จำนวนจังหวัดที่รายงานข้อมูลเบื้องต้น</td>
                        <td class="td-value">{sent_provinces_count:,} / 76</td>
                        <td class="td-progress">
                            <div style="font-size: 11px; text-align: right;">{round((sent_provinces_count/total_provinces)*100,2)}%</div>
                            <div class="progress-bg"><div class="progress-bar-cyan" style="width: {(sent_provinces_count/total_provinces)*100}%;"></div></div>
                        </td>
                        <td class="td-detail"><a href="/">เข้าระบบรายงาน</a></td>
                    </tr>
                    <tr>
                        <td class="td-label">สถานะการยืนยันและล็อคข้อมูล</td>
                        <td class="td-value">{locked_count:,} / 76</td>
                        <td class="td-progress">
                            <div style="font-size: 11px; text-align: right;">{locked_pct}%</div>
                            <div class="progress-bg"><div class="progress-bar-green" style="width: {locked_pct}%;"></div></div>
                        </td>
                        <td class="td-detail"></td>
                    </tr>
                    <tr>
                        <td class="td-label">สถานะการอัปโหลดไฟล์เอกสารรับรอง</td>
                        <td class="td-value">{uploaded_count:,} / 76</td>
                        <td class="td-progress">
                            <div style="font-size: 11px; text-align: right;">{uploaded_pct}%</div>
                            <div class="progress-bg"><div class="progress-bar-cyan" style="width: {uploaded_pct}%;"></div></div>
                        </td>
                        <td class="td-detail"></td>
                    </tr>
                    <tr>
                        <td class="td-label" style="font-weight: bold;">จำนวนโรงเรียนที่ร่วมสอบทั้งหมด</td>
                        <td class="td-value" colspan="2" style="text-align: left; color: #28a745;">{total_schools:,} แห่ง</td>
                        <td class="td-detail"></td>
                    </tr>
                </table>
            </div>

            <!-- Section 2: Split Cards -->
            <div class="grid-2">
                <!-- RT Card -->
                <div class="section-card card-rt">
                    <div class="section-header">📢 สรุปข้อมูลการจัดสอบ ป.1 (RT)</div>
                    
                    <div class="sub-header">ข้อมูลภาพรวมจังหวัด</div>
                    <div class="data-row"><span>จำนวนจังหวัดที่ส่งยอด ป.1</span><span class="val-num">{sent_provinces_count:,}</span></div>
                    
                    <div class="sub-header">ข้อมูลโรงเรียน</div>
                    <div class="data-row"><span>จำนวนโรงเรียนที่มีเด็กสอบ ป.1</span><span class="val-num">{total_schools:,}</span></div>
                    
                    <div class="sub-header">ข้อมูลนักเรียน</div>
                    <div class="data-row"><span>จำนวนนักเรียนปกติ</span><span class="val-num">{rt_normal:,}</span></div>
                    <div class="data-row"><span>จำนวนนักเรียนพิเศษ (เรียนร่วม)</span><span class="val-num">{rt_special:,}</span></div>
                    <div class="data-row" style="font-weight: bold; background-color: #fdf3f4;"><span>จำนวนนักเรียนทั้งสิ้น</span><span class="val-num">{rt_total:,}</span></div>
                </div>

                <!-- NT Card -->
                <div class="section-card card-nt">
                    <div class="section-header">📢 สรุปข้อมูลการจัดสอบ ป.3 (NT)</div>
                    
                    <div class="sub-header">ข้อมูลภาพรวมจังหวัด</div>
                    <div class="data-row"><span>จำนวนจังหวัดที่ส่งยอด ป.3</span><span class="val-num">{sent_provinces_count:,}</span></div>
                    
                    <div class="sub-header">ข้อมูลโรงเรียน</div>
                    <div class="data-row"><span>จำนวนโรงเรียนที่มีเด็กสอบ ป.3</span><span class="val-num">{total_schools:,}</span></div>
                    
                    <div class="sub-header">ข้อมูลนักเรียน</div>
                    <div class="data-row"><span>จำนวนนักเรียนปกติ</span><span class="val-num">{nt_normal:,}</span></div>
                    <div class="data-row"><span>จำนวนนักเรียนพิเศษ (เรียนร่วม)</span><span class="val-num">{nt_special:,}</span></div>
                    <div class="data-row" style="font-weight: bold; background-color: #f3fcf5;"><span>จำนวนนักเรียนทั้งสิ้น</span><span class="val-num">{nt_total:,}</span></div>
                </div>
            </div>
            
            <div style="text-align: center; margin-top: 20px;">
                <a href="/dashboard" style="color: #666; text-decoration: none; font-size: 14px;">← กลับไปหน้าเข้าสู่ระบบเจ้าหน้าที่</a>
            </div>
        </div>
    </body>
    </html>'''
    return HTMLResponse(content=html_content)