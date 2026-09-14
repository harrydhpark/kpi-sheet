const fs = require('fs');
const path = require('path');

const workspaceDir = "d:\\TV 유럽영업\\15. AX Task\\2026 AX 실행과제\\06. KPI Sheet";
const htmlPath = path.join(workspaceDir, "index.html");
const metricsPath = path.join(workspaceDir, "build_metrics.json");

const expectedRegions = [
  'LGEAG', 'LGEBN', 'LGECK', 'LGEDG', 'Swiss', 'LGEES', 
  'LGEFS', 'LGEHS', 'LGEIS', 'LGELA', 'LGEMK', 'LGEPL', 
  'LGEPT', 'LGERO', 'LGESW', 'LGEUK', 'EU', 'LGEAK', 
  'LGERA', 'LGEUR', 'CIS'
];

console.log("=======================================================");
console.log(" [DevOps Agent] 대시보드 빌드 무결성 검증 ");
console.log("=======================================================");

if (!fs.existsSync(htmlPath)) {
  console.error(`[FAIL] index.html 파일이 존재하지 않습니다: ${htmlPath}`);
  process.exit(1);
}

const stats = fs.statSync(htmlPath);
const sizeMB = stats.size / (1024 * 1024);
console.log(`* index.html 크기: ${sizeMB.toFixed(2)} MB`);

// Check size bounds (expected ~6.5MB - 7.5MB)
const sizeValid = sizeMB >= 6.0 && sizeMB <= 8.5;
if (!sizeValid) {
  console.warn(`[WARN] 파일 크기가 예상 범위를 벗어났습니다. (기준: 6.0MB ~ 8.5MB)`);
} else {
  console.log(`* 파일 크기 정합성: [PASS]`);
}

// Read sample of index.html to verify critical entities
console.log(`* 21개 법인/지점 데이터 포함 여부 검사 중...`);
const content = fs.readFileSync(htmlPath, 'utf8');

const missingRegions = [];
expectedRegions.forEach(reg => {
  // Check if region key exists in JSON payload
  if (!content.includes(`"${reg}"`) && !content.includes(`'${reg}'`)) {
    missingRegions.push(reg);
  }
});

let entityStatus = "PASS";
if (missingRegions.length > 0) {
  entityStatus = "FAIL";
  console.error(`[FAIL] 누락된 법인/지점이 있습니다: ${missingRegions.join(', ')}`);
} else {
  console.log(`* 21개 전 법인/지점 데이터 바인딩 확인: [PASS] (21/21)`);
}

// UI 필수 컨테이너 검사 (3대 테이블: PSI, P&L, Competitor)
const hasPsiTable = content.includes('id="table-psi"');
const hasPnlTable = content.includes('id="table-pnl"');
const hasCompTable = content.includes('id="table-comp"');
const hasSidebar = content.includes('EUROPE/CIS TV PORTAL');

console.log(`* 필수 UI 구성요소:`);
console.log(`  - PSI 테이블 (table-psi): ${hasPsiTable ? 'OK' : 'MISSING'}`);
console.log(`  - P&L 테이블 (table-pnl): ${hasPnlTable ? 'OK' : 'MISSING'}`);
console.log(`  - 경쟁지표 테이블 (table-comp): ${hasCompTable ? 'OK' : 'MISSING'}`);
console.log(`  - 사이드바/포털 타이틀: ${hasSidebar ? 'OK' : 'MISSING'}`);

const tablesOk = hasPsiTable && hasPnlTable && hasCompTable;
const overallStatus = (sizeValid && entityStatus === "PASS" && tablesOk && hasSidebar) ? "SUCCESS" : "WARNING";


const metrics = {
  timestamp: new Date().toISOString(),
  status: overallStatus,
  file_size_mb: parseFloat(sizeMB.toFixed(2)),
  regions_verified: expectedRegions.length - missingRegions.length,
  missing_regions: missingRegions,
  ui_checks: {
    psi_table: hasPsiTable,
    pnl_table: hasPnlTable,
    comp_table: hasCompTable,
    portal_sidebar: hasSidebar
  },

  preview_url: "http://localhost:4000",
  deploy_targets: ["Firebase Hosting", "LGE Internal GitLab (mod.lge.com)"]
};

fs.writeFileSync(metricsPath, JSON.stringify(metrics, null, 2), 'utf8');
console.log(`\n* 빌드 검증 메트릭 저장 완료: ${metricsPath}`);
console.log(`* 최종 빌드 검증 결과: [${overallStatus}]`);
console.log("=======================================================\n");

process.exit(overallStatus === "SUCCESS" ? 0 : 1);
