import pyxlsb
import os
import tempfile
import time
import json
import shutil
import zipfile
import win32com.client
from collections import defaultdict

WORKSPACE_DIR = r"d:\TV 유럽영업\15. AX Task\2026 AX 실행과제\06. KPI Sheet"
BACKUP_WORKBOOK = os.path.join(WORKSPACE_DIR, "backup", "26년_유럽+CIS_KPI_2026_pre_gfk07_20260922_095102.xlsx")
TARGET_WORKBOOK = os.path.join(WORKSPACE_DIR, "26년_유럽+CIS_KPI_2026.xlsx")
GFK_LOCAL = os.path.join(tempfile.gettempdir(), "gfk_local.xlsb")
JSON_OUTPUT = os.path.join(tempfile.gettempdir(), "gfk_july_extracted.json")

SUB_CODE_MAP = {
    'LGEAG': 'AG',
    'LGEBN': 'BN',
    'LGECK': 'CZ',
    'LGEDG': 'DG',
    'Swiss': 'DG_swiss',
    'LGEES': 'ES',
    'LGEFS': 'FS',
    'LGEHS': 'HS',
    'LGEIS': 'IS',
    'LGEMK': 'MK',
    'LGEPT': 'PT',
    'LGERO': 'RO',
    'LGESW': 'SW',
    'LGEUK': 'UK',
    'LGEAK': 'AK'
}
REV_MAP = {v: k for k, v in SUB_CODE_MAP.items()}

ROW_MAP = {
    'EU':     {'LG': 5,   'SS': 6,   'Sony': 7,   'TCL': 8,   'Hisense': 9,   '계': 11},
    'LGEAG':  {'LG': 12,  'SS': 13,  'Sony': 14,  'TCL': 15,  'Hisense': 16,  '계': 18},
    'LGEBN':  {'LG': 19,  'SS': 20,  'Sony': 21,  'TCL': 22,  'Hisense': 23,  '계': 25},
    'LGECK':  {'LG': 26,  'SS': 27,  'Sony': 28,  'TCL': 29,  'Hisense': 30,  '계': 32},
    'LGEDG':  {'LG': 33,  'SS': 34,  'Sony': 35,  'TCL': 36,  'Hisense': 37,  '계': 39},
    'Swiss':  {'LG': 40,  'SS': 41,  'Sony': 42,  'TCL': 43,  'Hisense': 44,  '계': 46},
    'LGEES':  {'LG': 47,  'SS': 48,  'Sony': 49,  'TCL': 50,  'Hisense': 51,  '계': 53},
    'LGEFS':  {'LG': 54,  'SS': 55,  'Sony': 56,  'TCL': 57,  'Hisense': 58,  '계': 60},
    'LGEHS':  {'LG': 61,  'SS': 62,  'Sony': 63,  'TCL': 64,  'Hisense': 65,  '계': 67},
    'LGEIS':  {'LG': 68,  'SS': 69,  'Sony': 70,  'TCL': 71,  'Hisense': 72,  '계': 74},
    'LGEMK':  {'LG': 75,  'SS': 76,  'Sony': 77,  'TCL': 78,  'Hisense': 79,  '계': 81},
    'LGEPT':  {'LG': 82,  'SS': 83,  'Sony': 84,  'TCL': 85,  'Hisense': 86,  '계': 88},
    'LGERO':  {'LG': 89,  'SS': 90,  'Sony': 91,  'TCL': 92,  'Hisense': 93,  '계': 95},
    'LGESW':  {'LG': 96,  'SS': 97,  'Sony': 98,  'TCL': 99,  'Hisense': 100, '계': 102},
    'LGEUK':  {'LG': 103, 'SS': 104, 'Sony': 105, 'TCL': 106, 'Hisense': 107, '계': 109},
    'CIS':    {'LG': 110, 'SS': 111, 'Sony': 112, 'TCL': 113, 'Hisense': 114, '계': 116},
    'LGEAK':  {'LG': 117, 'SS': 118, 'Sony': 119, 'TCL': 120, 'Hisense': 121, '계': 123},
}

def map_brand(b_str):
    if "01. LG" in b_str or "LG" in b_str:
        return "LG"
    elif "02." in b_str or "삼성" in b_str:
        return "SS"
    elif "03." in b_str or "소니" in b_str:
        return "Sony"
    elif "04." in b_str or "TCL" in b_str:
        return "TCL"
    elif "05." in b_str or "하이센스" in b_str:
        return "Hisense"
    return "기타"

def step1_extract():
    print("=== Step 1: pyxlsb Streaming Extraction ===")
    t0 = time.time()
    results = defaultdict(lambda: defaultdict(lambda: {'usd': 0.0, 'qty': 0.0}))
    matched = 0
    with pyxlsb.open_workbook(GFK_LOCAL) as wb:
        with wb.get_sheet('Raw') as sheet:
            for row_idx, row in enumerate(sheet.rows()):
                if row_idx < 2:
                    continue
                try:
                    month_val = row[35].v
                    if month_val != 202607 and month_val != "202607" and month_val != 202607.0:
                        continue
                    if str(row[3].v or '').strip() != 'O':
                        continue
                    sub_code = str(row[4].v or '').strip()
                    if sub_code not in REV_MAP:
                        continue
                    matched += 1
                    kpi_sub = REV_MAP[sub_code]
                    b = map_brand(str(row[18].v or '').strip())
                    qty = float(row[39].v or 0.0)
                    usd = float(row[40].v or 0.0)
                    results[kpi_sub][b]['usd'] += usd
                    results[kpi_sub][b]['qty'] += qty
                    results[kpi_sub]['계']['usd'] += usd
                    results[kpi_sub]['계']['qty'] += qty
                except Exception:
                    continue
                    
    print(f"Extraction done in {time.time() - t0:.1f}s. Matched rows: {matched}")
    
    # Calculate EU
    results['EU'] = defaultdict(lambda: {'usd': 0.0, 'qty': 0.0})
    for s in SUB_CODE_MAP:
        if s != 'LGEAK':
            for b in ['LG', 'SS', 'Sony', 'TCL', 'Hisense', '계']:
                results['EU'][b]['usd'] += results[s][b]['usd']
                results['EU'][b]['qty'] += results[s][b]['qty']
                
    results['CIS'] = results['LGEAK']
    
    # Save to JSON
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"Saved results to {JSON_OUTPUT}")
    return results

def step2_inject_and_recalc(data):
    print("=== Step 2: Injecting into Workbook & Recalculating ===")
    temp_dir = tempfile.gettempdir()
    temp_in = os.path.join(temp_dir, f"kpi_work_in_{os.getpid()}.xlsx")
    temp_out = os.path.join(temp_dir, f"kpi_work_out_{os.getpid()}.xlsx")
    
    # 1. Copy backup to temp_in
    print("Copying clean backup to local SSD...")
    shutil.copyfile(BACKUP_WORKBOOK, temp_in)
    
    # 2. Inject with openpyxl
    print("Injecting AL (Col 38) and BZ (Col 78) with openpyxl...")
    wb = openpyxl.load_workbook(temp_in)
    ws = wb['MI']
    for unit, b_dict in ROW_MAP.items():
        for b_name in ['LG', 'SS', 'Sony', 'TCL', 'Hisense', '계']:
            r = b_dict[b_name]
            u = data[unit][b_name]['usd']
            q = data[unit][b_name]['qty']
            ws.cell(row=r, column=38, value=u) # AL
            ws.cell(row=r, column=78, value=q) # BZ
    wb.save(temp_in)
    wb.close()
    print("Data injected successfully into MI sheet!")
    
    # 3. Excel COM for formula recalculation
    print("Recalculating formulas with Excel COM...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.Calculation = -4135 # manual
        wb_com = excel.Workbooks.Open(temp_in)
        excel.Calculation = -4105 # automatic
        wb_com.Application.Calculate()
        
        # Verify 3Q values
        ws_mi = wb_com.Sheets('MI')
        print("Verification from recalculated MI sheet:")
        for u in ['EU', 'LGEAG', 'LGEDG', 'CIS']:
            r_lg = ROW_MAP[u]['LG']
            r_tot = ROW_MAP[u]['계']
            print(f"  [{u}] LG USD(AL{r_lg})={ws_mi.Cells(r_lg,38).Value:,.0f} | QTY={ws_mi.Cells(r_lg,78).Value:,.0f} | ASP(DN{r_lg})={ws_mi.Cells(r_lg,118).Value:,.0f} | 3Q ASP(GD{r_lg})={ws_mi.Cells(r_lg,186).Value:,.0f}")
            
        wb_com.SaveCopyAs(temp_out)
        wb_com.Close(False)
    finally:
        excel.Quit()
        del excel
        time.sleep(2)
        
    # 4. Integrity check
    print("Verifying zip integrity of output...")
    with zipfile.ZipFile(temp_out, "r") as z:
        bad = z.testzip()
        if bad is not None:
            raise RuntimeError(f"Corrupt output zip: {bad}")
    print("Output integrity 100% OK!")
    
    # 5. Copy to NAS target
    print("Copying back to target NAS workbook...")
    shutil.copyfile(temp_out, TARGET_WORKBOOK)
    
    # Verify NAS target
    with zipfile.ZipFile(TARGET_WORKBOOK, "r") as z:
        bad = z.testzip()
        if bad is not None:
            raise RuntimeError(f"Corrupt NAS target: {bad}")
    print("Target workbook on NAS verified: 100% OK!")
    
    # Clean up temp files
    if os.path.exists(temp_in):
        os.remove(temp_in)
    if os.path.exists(temp_out):
        os.remove(temp_out)
    print("=== Pipeline Complete! ===")

if __name__ == "__main__":
    data = step1_extract()
    step2_inject_and_recalc(data)
