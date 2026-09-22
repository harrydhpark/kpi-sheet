import os
import sys
import json
import shutil
import tempfile
import time
import zipfile
import win32com.client

WORKSPACE_DIR = r"d:\TV 유럽영업\15. AX Task\2026 AX 실행과제\06. KPI Sheet"
BACKUP_WORKBOOK = os.path.join(WORKSPACE_DIR, "backup", "26년_유럽+CIS_KPI_2026_pre_gfk07_20260922_095102.xlsx")
TARGET_WORKBOOK = os.path.join(WORKSPACE_DIR, "26년_유럽+CIS_KPI_2026.xlsx")
JSON_INPUT = os.path.join(tempfile.gettempdir(), "gfk_july_extracted.json")

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

t0 = time.time()
print("=== 1. Loading July GfK JSON ===")
with open(JSON_INPUT, "r", encoding="utf-8") as f:
    data = json.load(f)
print(f"Loaded JSON. Units count: {len(data)}")

temp_dir = tempfile.gettempdir()
temp_in = os.path.join(temp_dir, f"kpi_com_in_{os.getpid()}.xlsx")
temp_out = os.path.join(temp_dir, f"kpi_com_out_{os.getpid()}.xlsx")

# 1. Copy clean backup to temp_in
print("=== 2. Copying backup to local SSD ===")
shutil.copyfile(BACKUP_WORKBOOK, temp_in)

# 2. Excel COM inject and calculate
print("=== 3. Injecting via Excel COM and Recalculating ===")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(temp_in)
    excel.Calculation = -4135 # manual calculation AFTER open!
    ws = wb.Sheets("MI")
    
    # Inject values
    for unit, b_dict in ROW_MAP.items():
        for b_name in ['LG', 'SS', 'Sony', 'TCL', 'Hisense', '계']:
            r = b_dict[b_name]
            u = data[unit][b_name]['usd']
            q = data[unit][b_name]['qty']
            ws.Cells(r, 38).Value = u # Col 38 = AL
            ws.Cells(r, 78).Value = q # Col 78 = BZ
            
    print("Cells injected. Triggering recalculation...")
    excel.Calculation = -4105 # automatic
    wb.Application.Calculate()
    
    # Verification
    print("\nRecalculation check:")
    for u in ['EU', 'LGEAG', 'LGEDG', 'CIS']:
        r_lg = ROW_MAP[u]['LG']
        al_val = ws.Cells(r_lg, 38).Value
        bz_val = ws.Cells(r_lg, 78).Value
        asp_val = ws.Cells(r_lg, 118).Value
        gd_val = ws.Cells(r_lg, 186).Value
        print(f"  [{u:5s}] AL(USD)={al_val:>14,.0f} | BZ(QTY)={bz_val:>9,.0f} | DN(ASP)={asp_val:>6.0f} | GD(3Q ASP)={gd_val:>6.0f}")
        
    wb.SaveCopyAs(temp_out)
    wb.Close(False)
finally:
    excel.Quit()
    del excel
    time.sleep(2)

# 3. Verify integrity
print("\n=== 4. Verifying ZIP integrity ===")
with zipfile.ZipFile(temp_out, "r") as z:
    bad = z.testzip()
    if bad is not None:
        raise RuntimeError(f"Corrupt output zip: {bad}")
print("Temp output integrity 100% OK!")

# 4. Copy to NAS target
print("=== 5. Copying to target workbook on NAS ===")
shutil.copyfile(temp_out, TARGET_WORKBOOK)

with zipfile.ZipFile(TARGET_WORKBOOK, "r") as z:
    bad = z.testzip()
    if bad is not None:
        raise RuntimeError(f"Corrupt NAS target: {bad}")
print("Target workbook on NAS verified: 100% OK!")

# Cleanup
if os.path.exists(temp_in):
    os.remove(temp_in)
if os.path.exists(temp_out):
    os.remove(temp_out)
    
print(f"\nAll operations finished successfully in {time.time() - t0:.2f}s!")
