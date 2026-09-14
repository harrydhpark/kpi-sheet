import os
import sys
import json
import datetime

WORKSPACE_DIR = r"d:\TV 유럽영업\15. AX Task\2026 AX 실행과제\06. KPI Sheet"
TEMP_DIR = r"C:\Users\harry.park\AppData\Local\Temp"
RAW_DATA_JSON = os.path.join(TEMP_DIR, "temp_raw_data.json")
OUTPUT_INSIGHTS = os.path.join(WORKSPACE_DIR, "executive_insights.json")

REGIONS = [
    'LGEAG', 'LGEBN', 'LGECK', 'LGEDG', 'Swiss', 'LGEES', 
    'LGEFS', 'LGEHS', 'LGEIS', 'LGELA', 'LGEMK', 'LGEPL', 
    'LGEPT', 'LGERO', 'LGESW', 'LGEUK', 'EU', 'LGEAK', 
    'LGERA', 'LGEUR', 'CIS'
]

# Sheet row numbers (1-based)
ROW_SELL_IN = 19
ROW_SELL_OUT = 22
ROW_WOS = 31
ROW_NET_SALES = 50
ROW_MARGIN_PROFIT = 56
ROW_OP = 59
ROW_MARGIN_RATE = 62
ROW_OPM = 65
ROW_NPI_SELLIN_RATIO = 69
ROW_NPI_SELLOUT_RATIO = 71

def safe_float(val, default=0.0):
    if val is None or val == "" or val == "#N/A" or val == "#DIV/0!" or val == "#VALUE!":
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default

def analyze_kpi():
    print("=======================================================")
    print(" [BI Analyst Agent] KPI 실적 다차원 분석 및 이상치 감지 ")
    print("=======================================================")
    
    if not os.path.exists(RAW_DATA_JSON):
        print(f"[ERROR] 추출된 KPI 데이터({RAW_DATA_JSON})를 찾을 수 없습니다.")
        print("먼저 DevOps Agent를 통해 대시보드 빌드 또는 데이터 추출을 실행하십시오.")
        return None
        
    with open(RAW_DATA_JSON, "r", encoding="utf-8-sig") as f:
        raw_data = json.load(f)
        
    insights = {
        "timestamp": datetime.datetime.now().isoformat(),
        "summary": {},
        "alerts": [],
        "regional_rankings": {
            "top_sales": [],
            "top_op": [],
            "risk_wos": []
        },
        "strategic_recommendations": []
    }
    
    regional_metrics = {}

    
    for reg in REGIONS:
        if reg not in raw_data:
            continue
            
        y26_raw = raw_data[reg].get("y26", {})
        y25_raw = raw_data[reg].get("y25", {})
        
        y26_matrix = y26_raw.get("value", y26_raw) if isinstance(y26_raw, dict) else y26_raw
        y25_matrix = y25_raw.get("value", y25_raw) if isinstance(y25_raw, dict) else y25_raw

        
        # Row data extraction
        # idx = row - 19
        def get_val(matrix, row_num, col_idx):
            r_idx = row_num - 19
            if 0 <= r_idx < len(matrix):
                row = matrix[r_idx]
                if 0 <= col_idx < len(row):
                    return safe_float(row[col_idx])
            return 0.0

        # Col 18 is typically YTD / H1 depending on model; Col 17 is 1H
        # Let's inspect Net Sales (Col 17=H1, Col 18=Jul YTD)
        net_sales_26 = get_val(y26_matrix, ROW_NET_SALES, 17)
        net_sales_25 = get_val(y25_matrix, ROW_NET_SALES, 17)
        op_26 = get_val(y26_matrix, ROW_OP, 17)
        op_25 = get_val(y25_matrix, ROW_OP, 17)
        opm_26 = get_val(y26_matrix, ROW_OPM, 17)
        wos_current = get_val(y26_matrix, ROW_WOS, 14) # 최근 6월/7월말 WOS
        margin_rate = get_val(y26_matrix, ROW_MARGIN_RATE, 17)
        npi_sellin = get_val(y26_matrix, ROW_NPI_SELLIN_RATIO, 17)
        
        regional_metrics[reg] = {
            "net_sales_m": round(net_sales_26 / 1e6, 1) if net_sales_26 > 1e4 else round(net_sales_26, 1),
            "net_sales_yoy_diff_m": round((net_sales_26 - net_sales_25) / 1e6, 1) if net_sales_26 > 1e4 else round(net_sales_26 - net_sales_25, 1),
            "op_m": round(op_26 / 1e6, 1) if abs(op_26) > 1e4 else round(op_26, 1),
            "op_yoy_diff_m": round((op_26 - op_25) / 1e6, 1) if abs(op_26) > 1e4 else round(op_26 - op_25, 1),
            "opm_percent": round(opm_26 * 100, 1) if abs(opm_26) < 1.0 else round(opm_26, 1),
            "margin_rate_percent": round(margin_rate * 100, 1) if abs(margin_rate) < 1.0 else round(margin_rate, 1),
            "wos_weeks": round(wos_current, 1),
            "npi_mix_percent": round(npi_sellin * 100, 1) if npi_sellin < 1.0 else round(npi_sellin, 1)
        }
        
        # 이상치 감지 규칙 (Anomaly Detection)
        if reg not in ['EU', 'CIS']: # 개별 법인/지점 대상
            # 1) WOS 리스크
            if wos_current > 16.0:
                insights["alerts"].append({
                    "level": "WARNING",
                    "category": "INVENTORY",
                    "entity": reg,
                    "indicator": "WOS",
                    "value": round(wos_current, 1),
                    "message": f"[{reg}] 유통재고 WOS가 {wos_current:.1f}주로 과다 재고 기준(16주)을 초과함. EOL 클리어런스 프로모션 필요."
                })
            elif 0.0 < wos_current < 4.0:
                insights["alerts"].append({
                    "level": "CRITICAL",
                    "category": "INVENTORY",
                    "entity": reg,
                    "indicator": "WOS",
                    "value": round(wos_current, 1),
                    "message": f"[{reg}] 유통재고 WOS가 {wos_current:.1f}주로 안전재고(4주) 미달. 공급 조기 배정 필요."
                })
                
            # 2) 수익성 리스크
            if regional_metrics[reg]["opm_percent"] < 0:
                insights["alerts"].append({
                    "level": "CRITICAL",
                    "category": "PROFITABILITY",
                    "entity": reg,
                    "indicator": "OPM",
                    "value": regional_metrics[reg]["opm_percent"],
                    "message": f"[{reg}] 영업이익률 마이너스({regional_metrics[reg]['opm_percent']}%) 적자 전환. 판가 및 차감율 긴급 점검 필요."
                })

    # HQ 집계
    eu_summary = regional_metrics.get("EU", {})
    cis_summary = regional_metrics.get("CIS", {})
    
    insights["summary"] = {
        "europe_hq": eu_summary,
        "cis_hq": cis_summary,
        "total_active_entities": len(regional_metrics)
    }
    
    # 랭킹 산출 (법인 대상)
    subs_only = {k: v for k, v in regional_metrics.items() if k not in ['EU', 'CIS']}
    sorted_by_sales = sorted(subs_only.items(), key=lambda x: x[1]["net_sales_m"], reverse=True)
    sorted_by_op = sorted(subs_only.items(), key=lambda x: x[1]["op_m"], reverse=True)
    sorted_by_wos = sorted(subs_only.items(), key=lambda x: x[1]["wos_weeks"], reverse=True)
    
    insights["regional_rankings"]["top_sales"] = [{"entity": k, "sales_m": v["net_sales_m"]} for k, v in sorted_by_sales[:5]]
    insights["regional_rankings"]["top_op"] = [{"entity": k, "op_m": v["op_m"], "opm": v["opm_percent"]} for k, v in sorted_by_op[:5]]
    insights["regional_rankings"]["risk_wos"] = [{"entity": k, "wos": v["wos_weeks"]} for k, v in sorted_by_wos[:5]]
    
    # 전략적 권고안 도출
    insights["strategic_recommendations"] = [
        "신모델(Y26 OLED G6/C6) 믹스 개선을 통한 고마진 구조(한계이익률 29%대) 지속 유지",
        f"WOS 주의 법인({', '.join([a['entity'] for a in insights['alerts'] if a['category'] == 'INVENTORY'][:3]) or '없음'}) 대상 이월 구모델 집중 클리어런스 실행",
        "독일/영국 등 핵심 대형 법인의 Sell-out 가속을 위한 유통사 공동 마케팅 펀드 우선 배정"
    ]
    
    with open(OUTPUT_INSIGHTS, "w", encoding="utf-8") as f:
        json.dump(insights, f, indent=2, ensure_ascii=False)
        
    print(f"* 유럽 본부(EU HQ) 실적: Net Sales ${eu_summary.get('net_sales_m')}M | OP ${eu_summary.get('op_m')}M (OPM {eu_summary.get('opm_percent')}%)")
    print(f"* 감지된 이상치/경보(Alerts): 총 {len(insights['alerts'])}건")
    for alert in insights["alerts"]:
        print(f"  - [{alert['level']}] {alert['message']}")
    print(f"\n* 분석 인사이트 저장 완료: {OUTPUT_INSIGHTS}")
    print("=======================================================\n")
    return insights

if __name__ == "__main__":
    analyze_kpi()
