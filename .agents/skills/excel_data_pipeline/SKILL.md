---
name: excel_data_pipeline
description: Excel COM batch calculation and normalization pipeline for TV business KPIs.
---

# Excel Data Pipeline Skill

본 스킬은 21개 지사/법인의 대용량 KPI 데이터를 Excel COM 개체를 사용하여 안전하고 신속하게 연산 및 정규화하는 데이터 가공 기술을 정의합니다.

## 1. 주요 처리 흐름

1. **파일 복사 및 임시 공간 활용**:
   - 한국어가 포함된 경로명은 PowerShell/Excel COM에서 파일 오프닝 시 인코딩 오류가 발생할 수 있습니다. 반드시 ASCII로 구성된 Windows Temp 폴더(`C:\Users\harry.park\AppData\Local\Temp\temp_input.xlsx`)로 파일을 임시 복사한 후 구동합니다.

2. **Excel COM 읽기전용 오픈**:
   - 백그라운드 자동화 시 파일이 잠기지 않도록 `ReadOnly` 플래그를 `$true`로 지정하여 엽니다:
     ```powershell
     $wb = $excel.Workbooks.Open($tempInput, 0, $true)
     ```
   - 경고 메시지 창이 나타나지 않도록 `DisplayAlerts = $false`, `AskToUpdateLinks = $false` 설정을 필수로 삽입합니다.

3. **지사 루프 순회 및 시트단위 부분 재계산**:
   - 통합 문서 전체 재계산(`$excel.Calculate()`)은 수십개의 피벗 테이블과 외부 연결 때문에 락이 걸리거나 오랜 시간이 걸립니다.
   - 반드시 H16 셀값 변경 후 해당 타겟 시트 2개만 명시하여 부분 계산(`$sheet.Calculate()`)합니다:
     ```powershell
     $sheet26.Range("H16").Value2 = "Swiss"
     $sheet26.Calculate()
     ```

4. **Range 일괄 추출 및 Jagged 배열 변환**:
   - 데이터는 `Range.Value2` 속성을 통해 2D multidimensional array (`System.Object[,]`) 형태로 한 번에 읽습니다.
   - 2026년 시트는 `$sheet26.Range("A19:BO293").Value2` 범위(75행에서 275행으로 확장), 2025년 시트는 `$sheet25.Range("A19:AZ293").Value2` 범위를 지정하여 일괄 추출해야 합니다.
   - 이를 PowerShell의 Jagged array (`System.Object[][]`) 형태로 루프 변환하여 JSON 직렬화에 적합한 데이터 구조로 구성합니다. 개별 셀 접근 시 `.GetValue(r, c)` 함수를 호출하여 에러를 예방합니다.

5. **2025년/2024년 데이터 분기 및 반기 취합**:
   - 엑셀 내 2025년 시트는 13개 컬럼(12개월 + TTL) 구조입니다. 대시보드 통일을 위해 상반기(H1), 1Q, 2Q, 3Q, 4Q 값을 메모리 상에서 합산 계산합니다.
   - **2025년 시트(KPI(25년)) 행 인덱스 매핑**: 2025년 시트는 2026년 시트와 레이아웃 번호가 다르므로 데이터 수집 시 반드시 신모델 OLED는 **Row 195**, 신모델 QNED는 **Row 207**, 구모델 OLED는 **Row 242**, 구모델 QNED는 **Row 254**의 데이터를 각각 참조해 합산해야 정확한 전년 실적이 산출됩니다.
   - **물량(n), 금액(m) 타입**: 분기 합산 및 상반기/연간 누적 계산. (단, 전시수량(행 73~78)은 예외적으로 재고/지수와 같이 최종 월 시점값을 적용합니다)
   - **재고/WOS(w) 및 지수(i) 타입**: 해당 분기/반기의 최종 월 기말 값 적용 (예: 1Q는 3월, 2Q/상반기는 6월, TTL은 12월 값).
   - **비율(p) 타입**: 단순 합산이 불가능하므로, 취합된 분자 행과 분모 행 데이터를 사용하여 분기/반기 시점 값을 재산산 처리합니다 (예: 신모델 비중 = 1Q 신모델 수량 / 1Q 전체 수량).
   - **지표 유형(Row Type) 분류 규칙**:
     - 행 이름에 "수량" 단어가 들어간 지표(예: `TV수량(제상품 매출)`)는 "매출" 단어가 섞여 있어도 수량(`n`)으로 정확히 분류될 수 있도록, 분류 조건문에서 "수량" 키워드를 "매출" 키워드보다 먼저 체크하도록 유형 탐색 순서를 구성합니다.

## 2. 정합성 검증 기준값 (Spot Checks)
- **Swiss (2026)**: Sell-in `66,022`, Sell-out `76,763`, Gross 매출 `79,384,983.69` ($79.4M), MS `21.87%` (21.9%)
- **LGERO (2025)**: Sell-in `211,672`, Sell-out `204,031.6`
