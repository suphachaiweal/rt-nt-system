@app.get("/export")
async def export_data(key: str = ""):
    if key != "DLA2569":
        return HTMLResponse("<script>alert('รหัสผ่านไม่ถูกต้อง! ระบบสงวนสิทธิ์เฉพาะเจ้าหน้าที่กรมเท่านั้น'); window.history.back();</script>")
        
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT province, dla_name, school_name, rt_student_count, nt_student_count, rt_special_json, nt_special_json FROM school_data ORDER BY province, dla_name, school_name", conn)
    
    if df.empty:
        conn.close()
        return HTMLResponse("<h2>ยังไม่มีข้อมูลให้ดาวน์โหลด</h2><a href='/dashboard'>กลับไปหน้าตรวจสอบ</a>")

    # คำนวณงบประมาณ
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
    file_path = "export_rt_nt_2569_calculated.xlsx"
    
    # --- เตรียมข้อมูลเด็กพิเศษ (Sheet 2) ---
    raw_rows = []
    def extract_sp_total(json_data):
        if not json_data or json_data == '{}': return 0
        try:
            data = json.loads(json_data) if isinstance(json_data, str) else json_data
            return sum(data.values())
        except: return 0
        
    def extract_sp_types(json_data):
        if not json_data or json_data == '{}': return {}
        try:
            return json.loads(json_data) if isinstance(json_data, str) else json_data
        except: return {}

    for index, row in df.iterrows():
        rt_sp = extract_sp_total(row['rt_special_json'])
        nt_sp = extract_sp_total(row['nt_special_json'])
        rt_normal = row['rt_student_count'] - rt_sp
        nt_normal = row['nt_student_count'] - nt_sp
        rt_types = extract_sp_types(row['rt_special_json'])
        nt_types = extract_sp_types(row['nt_special_json'])
        
        raw_rows.append({
            'จังหวัด': row['province'],
            'อปท.': row['dla_name'],
            'โรงเรียน': row['school_name'],
            'ป.1 (รวม)': row['rt_student_count'],
            'ป.1 (ปกติ)': rt_normal,
            'ป.1 (พิเศษ)': rt_sp,
            'ป.1 (ทางการเห็น)': rt_types.get('t1', 0),
            'ป.1 (ทางการได้ยิน)': rt_types.get('t2', 0),
            'ป.1 (สติปัญญา)': rt_types.get('t3', 0),
            'ป.1 (ร่างกาย)': rt_types.get('t4', 0),
            'ป.1 (LD)': rt_types.get('t5', 0),
            'ป.1 (พูดและภาษา)': rt_types.get('t6', 0),
            'ป.1 (พฤติกรรม)': rt_types.get('t7', 0),
            'ป.1 (ออทิสติก)': rt_types.get('t8', 0),
            'ป.1 (พิการซ้อน)': rt_types.get('t9', 0),
            'ป.3 (รวม)': row['nt_student_count'],
            'ป.3 (ปกติ)': nt_normal,
            'ป.3 (พิเศษ)': nt_sp,
            'ป.3 (ทางการเห็น)': nt_types.get('t1', 0),
            'ป.3 (ทางการได้ยิน)': nt_types.get('t2', 0),
            'ป.3 (สติปัญญา)': nt_types.get('t3', 0),
            'ป.3 (ร่างกาย)': nt_types.get('t4', 0),
            'ป.3 (LD)': nt_types.get('t5', 0),
            'ป.3 (พูดและภาษา)': nt_types.get('t6', 0),
            'ป.3 (พฤติกรรม)': nt_types.get('t7', 0),
            'ป.3 (ออทิสติก)': nt_types.get('t8', 0),
            'ป.3 (พิการซ้อน)': nt_types.get('t9', 0),
        })
    df_raw = pd.DataFrame(raw_rows)
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # Sheet 1: งบประมาณ (แบบฟอร์มเดิม)
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
        worksheet['A3'] = 'โครงการประเมินคุณภาพนักเรียนระดับการศึกษาภาคบังคับ ปีการศึกษา 2569'
        worksheet['A3'].font = Font(bold=True, size=12)
        worksheet['A3'].alignment = Alignment(horizontal='center', vertical='center')

        worksheet.merge_cells('A4:A5'); worksheet['A4'] = 'ลำดับ'
        worksheet.merge_cells('B4:B5'); worksheet['B4'] = 'จังหวัด/อปท./โรงเรียน'
        worksheet.merge_cells('G4:G5'); worksheet['G4'] = 'งบ อปท.\n(RT 1,000 / NT 1,000)'
        worksheet.merge_cells('H4:H5'); worksheet['H4'] = 'งบ สถจ.\n(RT 10,000 / NT 10,000)'
        worksheet.merge_cells('I4:I5'); worksheet['I4'] = 'รวมทั้งสิ้น\n(บาท)'
        
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
            
        # Sheet 2: ฐานข้อมูลเด็กพิเศษ
        df_raw.to_excel(writer, index=False, sheet_name='ข้อมูลเด็กพิเศษ (ดิบ)')
        ws_raw = writer.sheets['ข้อมูลเด็กพิเศษ (ดิบ)']
        for col in ws_raw.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try: 
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            adjusted_width = (max_length + 2)
            ws_raw.column_dimensions[column].width = adjusted_width

    return FileResponse(file_path, filename="สรุปงบประมาณ_RT_NT_2569.xlsx")