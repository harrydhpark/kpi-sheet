import os
import sys
import openpyxl

WORKSPACE_DIR = r"d:\TV 유럽영업\15. AX Task\2026 AX 실행과제\06. KPI Sheet"
BACKUP_FILE = os.path.join(WORKSPACE_DIR, "backup", "26년_유럽+CIS_KPI_2026_backup_20260914.xlsx")
CURRENT_FILE = os.path.join(WORKSPACE_DIR, "26년_유럽+CIS_KPI_2026.xlsx")

def compare_sheets():
    print("=======================================================")
    print(" [DataOps / BI Analyst] 백업본 vs 현재 KPI 엑셀 파일 정밀 비교")
    print("=======================================================")
    
    if not os.path.exists(BACKUP_FILE):
        print(f"[ERROR] 백업 파일 없음: {BACKUP_FILE}")
        return
    if not os.path.exists(CURRENT_FILE):
        print(f"[ERROR] 현재 파일 없음: {CURRENT_FILE}")
        return
        
    print(f"* 백업 파일: {os.path.basename(BACKUP_FILE)} ({os.path.getsize(BACKUP_FILE)/(1024*1024):.2f} MB)")
    print(f"* 현재 파일: {os.path.basename(CURRENT_FILE)} ({os.path.getsize(CURRENT_FILE)/(1024*1024):.2f} MB)")
    
    target_sheets = ["매출장Raw", "매출장Raw(신)", "매출장Raw(구)"]
    
    # openpyxl data_only=False 로 수식/값 비교
    print("\nLoading workbooks (data_only=False)...")
    wb_old = openpyxl.load_workbook(BACKUP_FILE, data_only=False)
    wb_new = openpyxl.load_workbook(CURRENT_FILE, data_only=False)
    
    months = ['01월', '02월', '03월', '04월', '05월', '06월', '07월', '08월', '09월', '10월', '11월', '12월']
    
    diff_summary = {}
    
    for sname in target_sheets:
        print(f"\n--- [{sname}] 시트 비교 중 ---")
        if sname not in wb_old.sheetnames or sname not in wb_new.sheetnames:
            print(f"시트 누락: {sname}")
            continue
            
        ws_old = wb_old[sname]
        ws_new = wb_new[sname]
        
        # Col range: AI:AT (35~46) for 매출장Raw, AG:AR (33~44) for 신/구
        start_col = 35 if sname == "매출장Raw" else 33
        
        changed_cells = []
        monthly_diff_count = {m: 0 for m in months}
        sample_changes = []
        
        # Row 5 to 130
        for r in range(5, 131):
            # 행 라벨 가져오기 (Col A, C, D, F 등)
            label = f"Row {r}"
            try:
                comp = ws_new.cell(r, 2).value or "" # 법인
                item = ws_new.cell(r, 4).value or ws_new.cell(r, 6).value or "" # 지표
                label = f"{comp} {item} (Row {r})".strip()
            except:
                pass
                
            for c_idx in range(12):
                col = start_col + c_idx
                m_name = months[c_idx]
                
                val_old = ws_old.cell(r, col).value
                val_new = ws_new.cell(r, col).value
                
                # 수식이나 값의 차이 검사
                diff = False
                if str(val_old) != str(val_new):
                    # 부동소수점 오차 체크
                    try:
                        f_old = float(val_old)
                        f_new = float(val_new)
                        if abs(f_old - f_new) > 1e-4:
                            diff = True
                    except:
                        diff = True
                        
                if diff:
                    monthly_diff_count[m_name] += 1
                    changed_cells.append((r, col, val_old, val_new))
                    if len(sample_changes) < 8:
                        sample_changes.append({
                            "label": label,
                            "month": m_name,
                            "old": val_old,
                            "new": val_new
                        })
                        
        diff_summary[sname] = {
            "total_changed_cells": len(changed_cells),
            "monthly_diff_count": monthly_diff_count,
            "samples": sample_changes
        }
        
        print(f"  * 변경된 총 셀 수: {len(changed_cells)}개")
        print(f"  * 월별 변경 셀 분포:")
        for m, cnt in monthly_diff_count.items():
            if cnt > 0:
                print(f"    - {m}: {cnt}개 셀 변동")
        if sample_changes:
            print("  * 주요 변경 샘플:")
            for sc in sample_changes[:4]:
                print(f"    - [{sc['label']} / {sc['month']}]: {sc['old']} -> {sc['new']}")

    wb_old.close()
    wb_new.close()
    # 인치사이즈별판매 시트 비교
    print("\n--- [인치사이즈별판매 정리] 시트 비교 중 ---")
    wb_old_data = openpyxl.load_workbook(BACKUP_FILE, data_only=True)
    wb_new_data = openpyxl.load_workbook(CURRENT_FILE, data_only=True)
    
    inch_sheets_old = [s for s in wb_old_data.sheetnames if "인치" in s]
    inch_sheets_new = [s for s in wb_new_data.sheetnames if "인치" in s]
    
    if inch_sheets_old and inch_sheets_new:
        ws_inch_old = wb_old_data[inch_sheets_old[0]]
        ws_inch_new = wb_new_data[inch_sheets_new[0]]
        
        inch_changed = 0
        inch_monthly_diff = {m: 0 for m in months}
        for r in range(5, min(ws_inch_new.max_row + 1, 2360)):
            for c_idx in range(12):
                c = 10 + c_idx
                vo = ws_inch_old.cell(r, c).value
                vn = ws_inch_new.cell(r, c).value
                if str(vo) != str(vn):
                    try:
                        if abs(float(vo or 0) - float(vn or 0)) > 1e-4:
                            inch_changed += 1
                            inch_monthly_diff[months[c_idx]] += 1
                    except:
                        inch_changed += 1
                        inch_monthly_diff[months[c_idx]] += 1
                        
        print(f"  * 인치 시트 ({inch_sheets_new[0]}): 총 {inch_changed}개 셀 변동")
        print("  * 월별 변동 분포:")
        for m, cnt in inch_monthly_diff.items():
            if cnt > 0:
                print(f"    - {m}: {cnt}개 셀 변동")
                
    # EU HQ (Row 101: TV 수량) 8~12월 수치 전후 비교
    print("\n--- [유럽 전체 (EU HQ) TV 출하 수량(Sell-in) 8~12월 비교] ---")
    ws_raw_old = wb_old_data["매출장Raw"]
    ws_raw_new = wb_new_data["매출장Raw"]
    # Row 101 is EU Total TV Qty (Col 35=01월 ... Col 42=08월 ... Col 46=12월)
    for c_idx in range(7, 12):
        col = 35 + c_idx
        m_name = months[c_idx]
        vo = ws_raw_old.cell(101, col).value
        vn = ws_raw_new.cell(101, col).value
        try:
            diff = float(vn or 0) - float(vo or 0)
            sign = "+" if diff > 0 else ""
            print(f"  * {m_name}: {vo:,.0f}대 -> {vn:,.0f}대 ({sign}{diff:,.0f}대, {((diff/float(vo))*100 if vo else 0):+.1f}%)")
        except:
            print(f"  * {m_name}: {vo} -> {vn}")
            
    wb_old_data.close()
    wb_new_data.close()
    print("\n=======================================================")
    print(" 비교 분석 완료")
    print("=======================================================")


if __name__ == "__main__":
    compare_sheets()
