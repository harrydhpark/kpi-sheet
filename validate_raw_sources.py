import os
import sys
import json
import datetime
import openpyxl

WORKSPACE_DIR = r"d:\TV 유럽영업\15. AX Task\2026 AX 실행과제\06. KPI Sheet"
RAW_DATA_DIR = os.path.join(WORKSPACE_DIR, "Raw data")
TARGET_WORKBOOK = os.path.join(WORKSPACE_DIR, "26년_유럽+CIS_KPI_2026.xlsx")
REPORT_JSON = os.path.join(WORKSPACE_DIR, "data_validation_report.json")

EXPECTED_REGIONS = [
    'LGEAG', 'LGEBN', 'LGECK', 'LGEDG', 'Swiss', 'LGEES', 
    'LGEFS', 'LGEHS', 'LGEIS', 'LGELA', 'LGEMK', 'LGEPL', 
    'LGEPT', 'LGERO', 'LGESW', 'LGEUK', 'EU', 'LGEAK', 
    'LGERA', 'LGEUR', 'CIS'
]

def check_raw_sources():
    sources = {
        "sales_ledger": {
            "name": "매출장",
            "file": "매출장.xlsb",
            "required": True,
            "desc": "Sell-in, Net Sales, 한계이익, 영업이익 (신/구모델 구분)"
        },
        "cpsi": {
            "name": "CPSI (Sell-out & WOS)",
            "file": "★ (Final) Y26 Consolidated CPSI.xlsx",
            "required": True,
            "desc": "Sell-out 수량, 유통재고, WOS"
        },
        "display": {
            "name": "Channel Inventory Display",
            "file": "Channel Inventory Display.xlsx",
            "required": True,
            "desc": "전시수량 및 인치/카테고리 자동 분류"
        },
        "inventory": {
            "name": "재고현황 (S&OP)",
            "file": "재고현황(DIO_LTI).xlsb",
            "required": False,
            "desc": "DIO, 장기재고율"
        },
        "mi_gfk": {
            "name": "MI GfK Raw",
            "file": "Global) GfK+NPD Monthly Raw_202501_202606.xlsb",
            "required": False,
            "desc": "시장판가(ASP), M/S, 전시지수"
        }
    }
    
    results = {}
    all_valid = True
    
    for key, info in sources.items():
        # 파일이름이 정확히 일치하거나 와일드카드 패턴 매칭
        filepath = os.path.join(RAW_DATA_DIR, info["file"])
        matched_file = None
        if os.path.exists(filepath):
            matched_file = filepath
        else:
            # 패턴 탐색 (예: date 포함 CPSI)
            if key == "cpsi":
                for f in os.listdir(RAW_DATA_DIR):
                    if "Consolidated CPSI" in f and not f.startswith("~$") and f.endswith(".xlsx"):
                        matched_file = os.path.join(RAW_DATA_DIR, f)
                        break
            elif key == "mi_gfk":
                for f in os.listdir(RAW_DATA_DIR):
                    if "GfK+NPD Monthly Raw" in f and not f.startswith("~$") and f.endswith(".xlsb"):
                        matched_file = os.path.join(RAW_DATA_DIR, f)
                        break
                        
        if matched_file and os.path.exists(matched_file):
            size_mb = os.path.getsize(matched_file) / (1024 * 1024)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(matched_file)).strftime('%Y-%m-%d %H:%M:%S')
            results[key] = {
                "status": "READY",
                "filename": os.path.basename(matched_file),
                "size_mb": round(size_mb, 2),
                "last_modified": mtime,
                "description": info["desc"]
            }
        else:
            status = "MISSING" if info["required"] else "NOT_PROVIDED"
            if info["required"]:
                all_valid = False
            results[key] = {
                "status": status,
                "filename": info["file"],
                "size_mb": 0,
                "last_modified": None,
                "description": info["desc"]
            }
            
    return results, all_valid

def check_target_workbook():
    if not os.path.exists(TARGET_WORKBOOK):
        return {
            "status": "MISSING",
            "error": "KPI 마스터 파일(26년_유럽+CIS_KPI_2026.xlsx)을 찾을 수 없습니다."
        }
        
    size_mb = os.path.getsize(TARGET_WORKBOOK) / (1024 * 1024)
    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(TARGET_WORKBOOK)).strftime('%Y-%m-%d %H:%M:%S')
    
    # openpyxl read-only 모드로 가볍게 시트 목록 및 validation 검사
    try:
        wb = openpyxl.load_workbook(TARGET_WORKBOOK, read_only=False, data_only=False)
        sheets = wb.sheetnames
        
        # H16 Data Validation 체크
        h16_valid = False
        if "KPI(26년)" in sheets:
            ws26 = wb["KPI(26년)"]
            for dv in ws26.data_validations.dataValidation:
                if "H16" in str(dv.sqref):
                    h16_valid = True
                    break
                    
        wb.close()
        
        return {
            "status": "HEALTHY",
            "filename": os.path.basename(TARGET_WORKBOOK),
            "size_mb": round(size_mb, 2),
            "last_modified": mtime,
            "sheets_count": len(sheets),
            "h16_dropdown_active": h16_valid
        }
    except Exception as e:
        return {
            "status": "WARNING",
            "filename": os.path.basename(TARGET_WORKBOOK),
            "size_mb": round(size_mb, 2),
            "last_modified": mtime,
            "error": str(e),
            "h16_dropdown_active": False
        }

def run_validation():
    print("=======================================================")
    print(" [DataOps Agent] 5대 원천 데이터 및 무결성 게이트 검증 ")
    print("=======================================================")
    
    source_results, sources_valid = check_raw_sources()
    wb_results = check_target_workbook()
    
    gate_status = "PASS" if sources_valid and wb_results["status"] == "HEALTHY" else "WARN"
    
    report = {
        "timestamp": datetime.datetime.now().isoformat(),
        "overall_status": gate_status,
        "raw_sources": source_results,
        "target_workbook": wb_results,
        "recommendation": "데이터 오버레이 파이프라인 진행 가능" if gate_status == "PASS" else "누락 파일 확인 또는 H16 드롭다운 복원 필요"
    }
    
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
        
    print(f"\n* 전체 상태 (Gate Status): [{gate_status}]")
    print(f"* 마스터 엑셀: {wb_results.get('filename')} ({wb_results.get('size_mb')} MB)")
    print(f"  - H16 셀 드롭다운 유효성: {'정상 활성화 (OK)' if wb_results.get('h16_dropdown_active') else '복원 필요 (Run restore_h16_dropdown.py)'}")
    print("\n* 5대 원천 데이터 파일 현황:")
    for k, v in source_results.items():
        print(f"  - [{v['status']:<12}] {v['filename']} ({v['size_mb']} MB) - {v['description']}")
        
    print(f"\n* 검증 리포트 저장 완료: {REPORT_JSON}")
    print("=======================================================\n")
    return gate_status

if __name__ == "__main__":
    status = run_validation()
    sys.exit(0 if status in ["PASS", "WARN"] else 1)
