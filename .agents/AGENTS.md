# KPI Sheet Europe — Project Rules

## Git Operations (LGE Internal GitLab)

이 프로젝트는 LGE 사내 GitLab(`mod.lge.com`)을 remote로 사용합니다.
Remote URL에 Personal Access Token이 포함되어 있으므로 아래 규칙을 반드시 따르십시오.

### 1. Windows Credential Manager 우회 필수

Windows 환경에서 `git push`를 실행할 때, Windows Credential Manager가 백그라운드에서 인증 팝업을 대기하며 **무한 hang**을 유발합니다.
URL에 이미 토큰이 포함되어 있으므로 credential helper 개입이 불필요합니다.

**git push/pull/fetch 실행 시 반드시 다음 환경변수를 설정하십시오:**

```powershell
$env:GIT_TERMINAL_PROMPT=0
$env:GIT_ASKPASS=''
git -c credential.helper='' push gitlab main --progress 2>&1
```

### 2. Git 명령어 PowerShell 호환

- PowerShell에서는 `&&` 연산자를 사용할 수 없습니다. 명령어 연결 시 `;` (세미콜론)을 사용하십시오.
- stderr 출력을 캡처하려면 `2>&1` 리다이렉션을 사용하십시오.

### 3. Remote 구성

- **Remote 이름**: `gitlab` (primary), `origin` (alias)
- **브랜치**: `main`
- **프로토콜**: HTTP (토큰 내장 URL)
- **http.postBuffer**: 524288000 (500MB, 대용량 파일 전송용)

### 4. 대용량 파일 및 index.html 수정 주의사항

- `index.html`이 약 6.9MB로 매우 큽니다.
- 최초 push 시 전체 object(~25MB)를 전송하므로 시간이 걸릴 수 있습니다.
- `git repack -a -d`를 사전 실행하면 전송 효율이 개선됩니다.
- 에이전트 도구(`multi_replace_file_content` 등)의 4MB 파일 크기 제한으로 인해 `index.html` 직접 편집이 실패할 경우, Node.js 헬퍼 스크립트(`fs.readFileSync` & `fs.writeFileSync`)를 생성하여 정밀 수정을 수행하십시오.
- `build_dashboard.js`와 `index.html` 두 파일 모두에 UI 및 로직 변경 사항을 항상 동기화하여 유지하십시오.

## 대시보드 UI/UX 및 명칭 표준 지침

### 1. 포털 및 헤더 타이틀 체계
- **대시보드 메인 배너 기본 타이틀**: `EUROPE/CIS TV Biz. KPI Monitoring Dashboard`
- **좌측 사이드바 서브타이틀**: `TV EUROPE/CIS TV PORTAL`
- **상단 툴바 표시**: `EUROPE/CIS TV PORTAL` | `[Page Indicator]`

### 2. 법인/지점 동적 타이틀 및 표기 표준
- **Europe HQ (`EU`)**: 상단 `EUROPE/CIS TV KPI Dashboard`, 배너 `유럽 TV Biz. KPI Monitoring Dashboard`
- **CIS HQ (`CIS`)**: 상단 `CIS TV KPI Dashboard`, 배너 `CIS TV Biz. KPI Monitoring Dashboard`
- **법인 (Subsidiaries)**:
  - Benelux (`LGEBN`): `베네룩스 법인`
  - Czech (`LGECK`): `체코 법인`
  - Germany (`LGEDG`): `독일 법인`
  - Spain (`LGEES`): `스페인 법인`
  - France (`LGEFS`): `프랑스 법인`
  - Italy (`LGEIS`): `이탈리아 법인`
  - Hungary (`LGEMK`): `헝가리 법인`
  - Poland (`LGEPL`): `폴란드 법인`
  - Portugal (`LGEPT`): `포르투갈 법인`
  - Sweden (`LGESW`): `스웨덴 법인`
  - United Kingdom (`LGEUK`): `영국 법인`
  - Kazakhstan (`LGEAK`): `카자흐스탄 법인`
  - Russia (`LGERA`): `러시아 법인`
- **지점 (Branches)**:
  - Austria (`LGEAG`): `오스트리아 지점`
  - Switzerland (`Swiss`): `스위스 지점`
  - Greece (`LGEHS`): `그리스 지점`
  - Latvia (`LGELA`): `Latvia (라트비아)` / `라트비아 지점` (사이드바 버튼: `Latvia`)
  - Portugal (`LGEPT`): `포르투갈 법인`
  - Romania (`LGERO`): `루마니아 지점`
  - Ukraine (`LGEUR`): `우크라이나 지점`

### 3. 요약 KPI 카드 증감 표시 규칙
- **Net Sales (`index 50`)**: `-$8.5M (▼ 0.7%)` (금액 차이 + YoY 퍼센트 표시)
- **영업이익 (`index 59`)**: `+$55.1M` (**증감 비율 % 표기 제외**, 절대 금액 증감량만 표시)
- **영업이익률 (`index 65`)**: `▲ 4.4%p` (포인트 차이 표시)

### 4. 데이터 테이블 고정 헤더 (Sticky Table) 스펙
- **`.table-container`**: `position: relative; overflow: auto; max-height: 650px;` (독립 박스 스크롤)
- **`th`**: `position: sticky; top: 0; background-color: #051c2c; color: #ffffff; z-index: 20;`
- **`.sticky-col`**: `position: sticky; left: 0; z-index: 10;` (행 배경색 `bg-white`, `bg-slate-50/50`, `hover:bg-slate-100/70`과 동적 매칭)
- **`.sticky-col-header`**: `position: sticky; left: 0; top: 0; background-color: #051c2c; color: #ffffff; z-index: 30;` (최상단-좌측 코너 셀 고정)

### 5. KPI Sheet 엑셀 데이터 갱신 시 H16 셀 드롭다운 목록 보존 필수
- openpyxl 라이브러리를 사용해 `26년_유럽+CIS_KPI_2026.xlsx`를 저장하면 엑셀의 Data Validation Extension이 삭제되는 현상이 있습니다.
- 따라서 매출장, CPSI 등 엑셀 데이터 파싱 오버레이 반영 후에는 **반드시 `restore_h16_dropdown.py`를 실행하여 `KPI(26년)` 및 `KPI(25년)` 시트 H16 셀의 법인/지점 선택 드롭다운 목록(Data Validation List: 21개 법인/지점/본부)을 복원**해야 합니다.

### 6. 매출장 신모델 및 구모델 분류 기준 표준
- **신모델 (New Model)**: `ReleaseYear == 'Y26'` (Col AC `출시 연도`)
- **구모델 (Old Model)**: `ReleaseYear`가 **`'Y26'`(신모델)과 `'X'`/`'x'`(불용/기타/비대상)를 제외한 나머지(Y25, Y24, Y23, Y22 등)** 유효 연도 모델
- **대상 시트 반영**:
  - `매출장Raw(신)`: 신모델 데이터 집계 오버레이 (EU/CIS 합계는 `SUMIF` 수식 유지)
  - `매출장Raw(구)`: 구모델 데이터 집계 오버레이 (EU/CIS 합계는 `SUMIF` 수식 유지)

### 7. 네트워크 드라이브(NAS) 대용량 엑셀 작업 시 로컬 버퍼(Local SSD Temp Buffer) 필수
- 작업 경로(`d:\TV 유럽영업\...`)는 사내 NAS 네트워크 스토리지입니다.
- 네트워크 상의 30MB+ 대용량 엑셀을 Excel COM으로 직접 열고 `$wb.SaveAs(...)`를 실행하면 파일 잠금 및 프로세스 무한 Hang이 발생합니다.
- 대용량 엑셀 수정 작업은 반드시 **`%TEMP%` 로컬 SSD 경로로 먼저 복사(`Copy-Item`)한 후 로컬에서 수정 및 `$wb.Save()` 저장하고, 완료 후 원본 경로로 덮어쓰기 복사**해야 합니다.

### 8. 경쟁 지표 (시장판가 ASP & M/S) 분기/상반기 집계 및 MI 시트 표준 구조
- **MI 시트 집계 컬럼 확장 스펙**:
  - `2026년 집계 (FQ~GJ열, Col 172~191)`: 2026 금액(1Q, 2Q, 상반기, 3Q, 4Q), 2026 수량, 2026 ASP, 2026 MS
  - `2025년 집계 (GK~HD열, Col 192~211)`: 2025 금액(1Q, 2Q, 상반기, 3Q, 4Q), 2025 수량, 2025 ASP, 2025 MS
  - **M/S 연산 공식**: `MS = Brand 금액 / Total 시장 금액` (`=IFERROR(FQ{brand}/FQ{market}, "")`)
  - **ASP 연산 공식**: `ASP = Brand 금액 / Brand 수량` (`=IFERROR(FQ{brand}/FV{brand}, "")`)
- **KPI(26년) 인덱스 및 VLOOKUP 참조 체계**:
  - **Row 11 (시장판가 ASP 인덱스)**:
    - 2026년: `R11(상반기)=184`, `Y11(1Q)=182`, `Z11(2Q)=183`, `AA11(3Q)=185`, `AB11(4Q)=186`
    - 2025년: `AK11(상반기)=204`, `AR11(1Q)=202`, `AS11(2Q)=203`, `AT11(3Q)=205`, `AU11(4Q)=206`
  - **Row 12 (M/S 인덱스)**:
    - 2026년: `R12(상반기)=189`, `Y12(1Q)=187`, `Z12(2Q)=188`, `AA12(3Q)=190`, `AB12(4Q)=191`
    - 2025년: `AK12(상반기)=209`, `AR12(1Q)=207`, `AS12(2Q)=208`, `AT12(3Q)=210`, `AU12(4Q)=211`
  - **Rows 83~88 (시장판가 ASP 수식)**: `=IFERROR(VLOOKUP($C{r}, MI!$B$5:$HD$123, {col}$11, 0), "")`
  - **Rows 89~93 (M/S 수식)**: `=IFERROR(VLOOKUP($C{r}, MI!$B$5:$HD$123, {col}$12, 0), "")`
- **대시보드(`build_dashboard.js`) 경쟁 지표 동기화 원칙**:
  - 경쟁 지표(83~93행)의 전년 실적(`row.y25`)은 `getAggregatesForRaw`의 단순 합산(ASP 판가 합산 및 M/S 0 처리 오류)을 우회하고, `KPI(26년)` 시트의 2025년 영역(`row.y25_prev`)을 1:1 직접 매핑하여 상반기, 1분기, 2분기 전년 및 전년비(YoY Diff)를 온전하게 산출해야 합니다.



