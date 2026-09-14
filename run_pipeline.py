import os
import sys
import json
import argparse
import subprocess
import datetime

WORKSPACE_DIR = r"d:\TV 유럽영업\15. AX Task\2026 AX 실행과제\06. KPI Sheet"
STATUS_JSON = os.path.join(WORKSPACE_DIR, "pipeline_status.json")
PYTHON_EXEC = r"C:\Users\harry.park\.gemini\antigravity\python_portable\python.exe"

def update_status(step, status, message=None):
    pipeline_state = {}
    if os.path.exists(STATUS_JSON):
        try:
            with open(STATUS_JSON, "r", encoding="utf-8") as f:
                pipeline_state = json.load(f)
        except Exception:
            pipeline_state = {}
            
    pipeline_state["last_updated"] = datetime.datetime.now().isoformat()
    if "steps" not in pipeline_state:
        pipeline_state["steps"] = {}
        
    pipeline_state["steps"][step] = {
        "status": status,
        "timestamp": datetime.datetime.now().isoformat(),
        "message": message
    }
    
    with open(STATUS_JSON, "w", encoding="utf-8") as f:
        json.dump(pipeline_state, f, indent=2, ensure_ascii=False)

def run_dataops_validation():
    print("\n>>> [Step 1 / DataOps Agent] 5대 원천 데이터 및 무결성 게이트 검증 시작...")
    update_status("dataops_validation", "RUNNING")
    cmd = [PYTHON_EXEC, os.path.join(WORKSPACE_DIR, "validate_raw_sources.py")]
    res = subprocess.run(cmd, cwd=WORKSPACE_DIR)
    if res.returncode == 0:
        update_status("dataops_validation", "COMPLETED", "원천 데이터 및 H16 드롭다운 검증 통과")
        print(">>> [DataOps Agent] 검증 통과 (SUCCESS)")
        return True
    else:
        update_status("dataops_validation", "FAILED", "데이터 검증 게이트 실패")
        print(">>> [DataOps Agent] 검증 실패 (FAILED)")
        return False

def run_devops_build():
    print("\n>>> [Step 2 / DevOps Agent] 대시보드 빌드 무결성 검증 시작...")
    update_status("devops_build", "RUNNING")
    
    # 1. Validate existing/compiled index.html
    cmd_validate = ["node", "validate_dashboard_build.js"]
    res_val = subprocess.run(cmd_validate, cwd=WORKSPACE_DIR)
    if res_val.returncode == 0:
        update_status("devops_build", "COMPLETED", "index.html 21개 법인 및 필수 테이블 무결성 검증 완료")
        print(">>> [DevOps Agent] 대시보드 빌드 검증 통과 (SUCCESS)")
        return True
    else:
        update_status("devops_build", "FAILED", "대시보드 빌드 검증 실패")
        print(">>> [DevOps Agent] 대시보드 빌드 검증 실패 (FAILED)")
        return False

def run_bi_analysis():
    print("\n>>> [Step 3 / BI Analyst Agent] KPI 실적 다차원 분석 및 이상치 감지 시작...")
    update_status("bi_analysis", "RUNNING")
    cmd = [PYTHON_EXEC, os.path.join(WORKSPACE_DIR, "analyze_kpi_insights.py")]
    res = subprocess.run(cmd, cwd=WORKSPACE_DIR)
    if res.returncode == 0:
        update_status("bi_analysis", "COMPLETED", "21개 법인 실적 진단 및 이상치 감지 리포트 생성 완료")
        print(">>> [BI Analyst Agent] 실적 진단 완료 (SUCCESS)")
        return True
    else:
        update_status("bi_analysis", "FAILED", "실적 진단 실패")
        print(">>> [BI Analyst Agent] 실적 진단 실패 (FAILED)")
        return False

def main():
    parser = argparse.ArgumentParser(description="KPI Sheet Master Orchestrator Pipeline Runner")
    parser.add_argument("--mode", choices=["validate", "build", "analyze", "full"], default="full",
                        help="실행할 파이프라인 모드: validate (DataOps), build (DevOps), analyze (BI Analyst), full (전체)")
    args = parser.parse_args()

    print("=================================================================")
    print(" [Master Orchestrator] KPI Sheet 멀티 에이전트 파이프라인 가동 ")
    print(f" 모드: {args.mode.upper()} | 시작 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=================================================================")

    if args.mode in ["validate", "full"]:
        if not run_dataops_validation() and args.mode == "full":
            print("\n[중단] DataOps 품질 게이트 실패로 후속 파이프라인 중단.")
            sys.exit(1)

    if args.mode in ["build", "full"]:
        if not run_devops_build() and args.mode == "full":
            print("\n[중단] DevOps 빌드 검증 실패로 후속 파이프라인 중단.")
            sys.exit(1)

    if args.mode in ["analyze", "full"]:
        if not run_bi_analysis() and args.mode == "full":
            print("\n[중단] BI Analyst 실적 진단 실패.")
            sys.exit(1)

    print("\n=================================================================")
    print(" [Master Orchestrator] 모든 요청 에이전트 파이프라인 정상 완료! ")
    print(f" 상태 저장 위치: {STATUS_JSON}")
    print("=================================================================\n")

if __name__ == "__main__":
    main()
