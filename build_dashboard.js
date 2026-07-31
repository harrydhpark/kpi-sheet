const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const workspaceDir = "d:\\TV 유럽영업\\15. AX Task\\2026 AX 실행과제\\06. KPI Sheet";
const originalFile = path.join(workspaceDir, "26년_유럽+CIS_KPI_260712 1.xlsx");

const tempDir = "C:\\Users\\harry.park\\AppData\\Local\\Temp";
const tempInput = path.join(tempDir, "temp_input.xlsx");
const rawDataJson = path.join(tempDir, "temp_raw_data.json");
const psScriptPath = path.join(tempDir, "extract.ps1");
const finalHtmlPath = path.join(workspaceDir, "index.html");

console.log("=== STEP 1: Excel COM Batch Data Extraction ===");

// Force kill Excel first to ensure no locks
try {
  console.log("Killing any existing Excel processes...");
  execSync('powershell -Command "Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force"', { stdio: 'ignore' });
} catch (e) {
  // Ignore
}

console.log("Copying original workbook to temp path...");
if (fs.existsSync(tempInput)) {
  fs.unlinkSync(tempInput);
}
fs.copyFileSync(originalFile, tempInput);
console.log("Copied to:", tempInput);

const regions = [
  'LGEAG', 'LGEBN', 'LGECK', 'LGEDG', 'Swiss', 'LGEES', 
  'LGEFS', 'LGEHS', 'LGEIS', 'LGELA', 'LGEMK', 'LGEPL', 
  'LGEPT', 'LGERO', 'LGESW', 'LGEUK', 'EU', 'LGEAK', 
  'LGERA', 'LGEUR', 'CIS'
];

const psScriptContent = `
$tempInput = '${tempInput}'
$rawDataJson = '${rawDataJson}'
$regions = @(${regions.map(r => `'${r}'`).join(', ')})

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

try {
    Write-Host "Opening temp input workbook in Excel COM (ReadOnly)..."
    $wb = $excel.Workbooks.Open($tempInput, 0, $true)
    
    $sheet26 = $wb.Sheets.Item("KPI(26년)")
    $sheet25 = $wb.Sheets.Item("KPI(25년)")
    
    $extracted = @{}
    
    foreach ($r in $regions) {
        Write-Host "Extracting data for region: $r..."
        
        $sheet26.Range("H16").Value2 = $r
        $sheet25.Range("H16").Value2 = $r
        
        $sheet26.Calculate()
        $sheet25.Calculate()
        
        # Read entire range as 2D array
        $arr26 = $sheet26.Range("A19:BO293").Value2
        $arr25 = $sheet25.Range("A19:AZ293").Value2
        
        # Convert 2D arrays to jagged arrays for JSON serialization using .GetValue()
        $jagged26 = New-Object System.Object[][] $arr26.GetLength(0)
        for ($i = 0; $i -lt $arr26.GetLength(0); $i++) {
            $rowArr = New-Object System.Object[] $arr26.GetLength(1)
            for ($j = 0; $j -lt $arr26.GetLength(1); $j++) {
                $rowArr[$j] = $arr26.GetValue($i + 1, $j + 1)
            }
            $jagged26[$i] = $rowArr
        }
        
        $jagged25 = New-Object System.Object[][] $arr25.GetLength(0)
        for ($i = 0; $i -lt $arr25.GetLength(0); $i++) {
            $rowArr = New-Object System.Object[] $arr25.GetLength(1)
            for ($j = 0; $j -lt $arr25.GetLength(1); $j++) {
                $rowArr[$j] = $arr25.GetValue($i + 1, $j + 1)
            }
            $jagged25[$i] = $rowArr
        }
        
        $extracted[$r] = @{
            "y26" = $jagged26
            "y25" = $jagged25
        }
    }
    
    Write-Host "Writing extracted data to JSON file..."
    if (Test-Path $rawDataJson) { Remove-Item $rawDataJson -Force }
    $extracted | ConvertTo-Json -Depth 5 | Out-File $rawDataJson -Encoding utf8
    Write-Host "Extraction completed successfully."
}
catch {
    Write-Error $_
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
`;

// Write with UTF-8 BOM
fs.writeFileSync(psScriptPath, '\ufeff' + psScriptContent, 'utf8');
console.log("Created PowerShell script with BOM at:", psScriptPath);

console.log("Executing PowerShell extraction script...");
try {
  const output = execSync(`powershell -ExecutionPolicy Bypass -File "${psScriptPath}"`, { encoding: 'utf8' });
  console.log("PowerShell Output:\n", output);
} catch (err) {
  console.error("PowerShell failed:", err.stdout || err.message);
  process.exit(1);
}

console.log("=== STEP 2: Parser and Expansion Engine ===");
let jsonStr = fs.readFileSync(rawDataJson, 'utf8');
if (jsonStr.startsWith('\ufeff')) {
  jsonStr = jsonStr.slice(1);
}
const rawData = JSON.parse(jsonStr);

// Helper to determine the type of a row based on labels and descriptor
function getRowType(labels) {
  const fullLabel = labels.filter(l => l !== null && l !== undefined).map(l => String(l).trim()).join(' ').toLowerCase();
  
  if (fullLabel.endsWith('신모델') || fullLabel.endsWith('구모델')) {
    return 'n';
  }
  if (fullLabel.includes('%') || fullLabel.includes('비중') || fullLabel.includes('m/s') || fullLabel.includes('가동율') || fullLabel.includes('차감율')) {
    return 'p';
  }
  if (fullLabel.includes('wos')) {
    return 'w';
  }
  if (fullLabel.includes('지수')) {
    return 'i';
  }
  if (fullLabel.includes('수량')) {
    return 'n';
  }
  if (fullLabel.includes('매출') || fullLabel.includes('이익') || fullLabel.includes('금액') || fullLabel.includes('asp')) {
    return 'm';
  }
  return 'n';
}

const processedData = {};

regions.forEach(region => {
  const data26 = rawData[region].y26.value;
  const data25 = rawData[region].y25.value;
  const rowsList = [];
  
  for (let idx = 0; idx < 275; idx++) {
    const row26 = data26[idx];
    const sheetRowNumber = idx + 19;
    let idx25 = idx;
    if (sheetRowNumber === 198) {
      idx25 = 176; // Row 195 in KPI(25년)
    } else if (sheetRowNumber === 213) {
      idx25 = 188; // Row 207 in KPI(25년)
    } else if (sheetRowNumber === 248) {
      idx25 = 223; // Row 242 in KPI(25년)
    } else if (sheetRowNumber === 260) {
      idx25 = 235; // Row 254 in KPI(25년)
    }
    const row25 = data25[idx25];
    const labels = [row26[6], row26[7], row26[8], row26[9], row26[10]];
    let type = getRowType(labels);
    
    // Custom label override for Row 47
    if (sheetRowNumber === 47) {
      labels[1] = "Net매출(제상품)";
    }
    
    // Custom label override for Rows 69 and 71
    if (sheetRowNumber === 69) {
      labels[1] = "신모델 Sell-in 비중(수량)";
      labels[2] = null;
      labels[3] = null;
      labels[4] = null;
    } else if (sheetRowNumber === 71) {
      labels[1] = "신모델 Sell-out 비중(수량)";
      labels[2] = null;
      labels[3] = null;
      labels[4] = null;
    }
    
    // Custom label override for Competitor Table (indices 73-93)
    if (sheetRowNumber >= 73 && sheetRowNumber <= 78) {
      for (let i = 4; i >= 0; i--) {
        if (labels[i] !== null && labels[i] !== undefined && String(labels[i]).trim() !== '') {
          labels[i] = "전시수량 (" + String(labels[i]).trim() + ")";
          break;
        }
      }
    } else if (sheetRowNumber >= 79 && sheetRowNumber <= 82) {
      for (let i = 4; i >= 0; i--) {
        if (labels[i] !== null && labels[i] !== undefined && String(labels[i]).trim() !== '') {
          labels[i] = "전시지수 (" + String(labels[i]).trim() + ")";
          break;
        }
      }
    } else if (sheetRowNumber >= 83 && sheetRowNumber <= 88) {
      for (let i = 4; i >= 0; i--) {
        if (labels[i] !== null && labels[i] !== undefined && String(labels[i]).trim() !== '') {
          labels[i] = "시장판가 (" + String(labels[i]).trim() + ")";
          break;
        }
      }
    } else if (sheetRowNumber >= 89 && sheetRowNumber <= 93) {
      for (let i = 4; i >= 0; i--) {
        if (labels[i] !== null && labels[i] !== undefined && String(labels[i]).trim() !== '') {
          labels[i] = "M/S (" + String(labels[i]).trim() + ")";
          break;
        }
      }
    }
    
    // Explicit row type overrides for P&L Table (indices 38-67)
    if (sheetRowNumber >= 44 && sheetRowNumber <= 46) {
      type = 'p'; // 차감 및 하위 항목
    } else if (sheetRowNumber >= 47 && sheetRowNumber <= 52) {
      type = 'm'; // Net매출 및 하위 항목
    } else if (sheetRowNumber >= 53 && sheetRowNumber <= 55) {
      type = 'm'; // N.ASP($) 및 하위 항목
    } else if (sheetRowNumber >= 56 && sheetRowNumber <= 61) {
      type = 'm'; // 한계이익(금액), 영업이익(금액) 및 하위 항목
    } else if (sheetRowNumber >= 62 && sheetRowNumber <= 67) {
      type = 'p'; // 한계이익(%), 영업이익(%) 및 하위 항목
    }
    
    // Explicit row type overrides for Competitor Table (indices 73-93)
    if (sheetRowNumber >= 79 && sheetRowNumber <= 82) {
      type = 'i'; // 전시지수
    } else if (sheetRowNumber >= 83 && sheetRowNumber <= 88) {
      type = 'm'; // 시장판가
    } else if (sheetRowNumber >= 89 && sheetRowNumber <= 93) {
      type = 'p'; // M/S
    }
    
    // Explicit row type overrides for Rows 69 and 71
    if (sheetRowNumber === 69 || sheetRowNumber === 71) {
      type = 'p';
    }
    
    // 2026 range values (L to AC, indices 11 to 28)
    const y26_vals = row26.slice(11, 29).map(v => (v === null || v === undefined || v === '') ? null : Number(v));
    // 2025 previous range values (AE to AV, indices 30 to 47)
    const y25_prev_vals = row26.slice(30, 48).map(v => (v === null || v === undefined || v === '') ? null : Number(v));
    // YoY values (AX to BO, indices 49 to 66)
    const y26_yoy_vals = row26.slice(49, 67).map(v => (v === null || v === undefined || v === '') ? null : Number(v) * 100);
    
    // 2025 raw data
    const y25_raw_vals = row25.slice(11, 24).map(v => (v === null || v === undefined || v === '') ? null : Number(v));
    const y24_raw_vals = row25.slice(25, 38).map(v => (v === null || v === undefined || v === '') ? null : Number(v));
    
    rowsList.push({
      index: sheetRowNumber,
      labels: labels,
      type: type,
      y26: y26_vals,
      y25_prev: y25_prev_vals,
      y26_yoy: y26_yoy_vals,
      y25_raw: y25_raw_vals,
      y24_raw: y24_raw_vals,
      y25: null,
      y24: null,
      y25_yoy: null
    });
  }
  
  rowsList.forEach(row => {
    row.y25 = getAggregatesForRaw(row.y25_raw, row.type, row.index);
    row.y24 = getAggregatesForRaw(row.y24_raw, row.type, row.index);
  });
  
  computeRatios(rowsList, 'y25');
  computeRatios(rowsList, 'y24');
  
  rowsList.forEach(row => {
    row.y25_yoy = new Array(18);
    for (let c = 0; c < 18; c++) {
      const v25 = row.y25[c];
      const v24 = row.y24[c];
      if (v24 && v24 !== 0 && v25 !== null) {
        row.y25_yoy[c] = (v25 / v24 - 1) * 100;
      } else {
        row.y25_yoy[c] = null;
      }
    }
  });
  
  // Sum Sell-out OLED and QNED virtual rows
  const row198 = rowsList.find(r => r.index === 198);
  const row248 = rowsList.find(r => r.index === 248);
  const row213 = rowsList.find(r => r.index === 213);
  const row260 = rowsList.find(r => r.index === 260);

  function sumRows(rowA, rowB, newIndex, newLabels, newType) {
    const sumRow = {
      index: newIndex,
      labels: newLabels,
      type: newType,
      y26: new Array(18).fill(null),
      y25_prev: new Array(18).fill(null),
      y26_yoy: new Array(18).fill(null),
      y25_raw: new Array(13).fill(null),
      y24_raw: new Array(13).fill(null),
      y25: new Array(18).fill(null),
      y24: new Array(18).fill(null),
      y25_yoy: new Array(18).fill(null)
    };
    
    const fields = ['y26', 'y25_prev', 'y25_raw', 'y24_raw', 'y25', 'y24'];
    fields.forEach(field => {
      const len = sumRow[field].length;
      for (let c = 0; c < len; c++) {
        const valA = rowA && rowA[field] ? rowA[field][c] : null;
        const valB = rowB && rowB[field] ? rowB[field][c] : null;
        if (valA !== null || valB !== null) {
          sumRow[field][c] = (valA || 0) + (valB || 0);
        }
      }
    });
    
    for (let c = 0; c < 18; c++) {
      const v26 = sumRow.y26[c];
      const v25p = sumRow.y25_prev[c];
      if (v25p && v25p !== 0 && v26 !== null) {
        sumRow.y26_yoy[c] = (v26 / v25p - 1) * 100;
      } else {
        sumRow.y26_yoy[c] = null;
      }
    }
    
    for (let c = 0; c < 18; c++) {
      const v25 = sumRow.y25[c];
      const v24 = sumRow.y24[c];
      if (v24 && v24 !== 0 && v25 !== null) {
        sumRow.y25_yoy[c] = (v25 / v24 - 1) * 100;
      } else {
        sumRow.y25_yoy[c] = null;
      }
    }
    
    return sumRow;
  }

  const idx24 = rowsList.findIndex(r => r.index === 24);
  if (idx24 !== -1) {
    const virtualRows = [];
    if (row198 || row248) {
      const oledSumRow = sumRows(row198, row248, 24.1, [null, null, "OLED", null, null], 'n');
      virtualRows.push(oledSumRow);
    }
    if (row213 || row260) {
      const qnedSumRow = sumRows(row213, row260, 24.2, [null, null, "QNED", null, null], 'n');
      virtualRows.push(qnedSumRow);
    }
    rowsList.splice(idx24 + 1, 0, ...virtualRows);
  }

  processedData[region] = rowsList;
});

function getAggregatesForRaw(vals, type, sheetRowNumber) {
  const result = new Array(18);
  for (let m = 0; m < 6; m++) result[m] = vals[m];
  for (let m = 6; m < 12; m++) result[m + 1] = vals[m];
  
  // Custom Stock-like aggregation for 전시수량 (rows 73 to 78)
  const isDisplayQty = sheetRowNumber !== undefined && sheetRowNumber >= 73 && sheetRowNumber <= 78;
  
  if (isDisplayQty) {
    result[6] = vals[5];   // 상반기 -> 6월
    result[13] = vals[2];  // 1Q -> 3월
    result[14] = vals[5];  // 2Q -> 6월
    result[15] = vals[8];  // 3Q -> 9월
    result[16] = vals[11]; // 4Q -> 12월
    result[17] = vals[11]; // TTL -> 12월
  } else if (type === 'n' || type === 'm') {
    const q1 = (vals[0] || 0) + (vals[1] || 0) + (vals[2] || 0);
    const q2 = (vals[3] || 0) + (vals[4] || 0) + (vals[5] || 0);
    const h1 = q1 + q2;
    const q3 = (vals[6] || 0) + (vals[7] || 0) + (vals[8] || 0);
    const q4 = (vals[9] || 0) + (vals[10] || 0) + (vals[11] || 0);
    result[6] = h1;
    result[13] = q1;
    result[14] = q2;
    result[15] = q3;
    result[16] = q4;
    result[17] = vals[12];
  } else if (type === 'w' || type === 'i') {
    result[6] = vals[5];
    result[13] = vals[2];
    result[14] = vals[5];
    result[15] = vals[8];
    result[16] = vals[11];
    result[17] = vals[12];
  } else {
    result[6] = 0;
    result[13] = 0;
    result[14] = 0;
    result[15] = 0;
    result[16] = 0;
    result[17] = vals[12];
  }
  return result;
}

function computeRatios(rowList, yearKey) {
  const getRow = (idx) => rowList.find(r => r.index === idx);
  
  const divRows = (numRowIdx, denRowIdx, targetRowIdx) => {
    const target = getRow(targetRowIdx);
    const num = getRow(numRowIdx);
    const den = getRow(denRowIdx);
    if (!target) return;
    for (let c = 0; c < 18; c++) {
      const nVal = (num && num[yearKey][c] !== null) ? num[yearKey][c] : 0;
      const dVal = (den && den[yearKey][c] !== null) ? den[yearKey][c] : 0;
      target[yearKey][c] = (dVal && dVal !== 0) ? (nVal / dVal) : null;
    }
  };
  
  divRows(25, 24, 26);
  divRows(22, 19, 69);
  divRows(25, 24, 71);
  
  const gross = getRow(38);
  const net = getRow(47);
  const target44 = getRow(44);
  if (target44 && gross && net) {
    for (let c = 0; c < 18; c++) {
      const g = gross[yearKey][c];
      const n = net[yearKey][c];
      target44[yearKey][c] = g ? (g - n) / g : 0;
    }
  }
  
  const grossNew = getRow(39);
  const netNew = getRow(48);
  const target45 = getRow(45);
  if (target45 && grossNew && netNew) {
    for (let c = 0; c < 18; c++) {
      const g = grossNew[yearKey][c];
      const n = netNew[yearKey][c];
      target45[yearKey][c] = g ? (g - n) / g : 0;
    }
  }
  
  const grossOld = getRow(40);
  const netOld = getRow(49);
  const target46 = getRow(46);
  if (target46 && grossOld && netOld) {
    for (let c = 0; c < 18; c++) {
      const g = grossOld[yearKey][c];
      const n = netOld[yearKey][c];
      target46[yearKey][c] = g ? (g - n) / g : 0;
    }
  }
  
  divRows(56, 50, 62);
  divRows(57, 51, 63);
  divRows(58, 52, 64);
  
  divRows(59, 50, 65);
  divRows(60, 51, 66);
  divRows(61, 52, 67);
  
  divRows(22, 19, 69);
  divRows(48, 47, 70);
  
  const target71 = getRow(71);
  const r26 = getRow(26);
  if (target71 && r26) {
    for (let c = 0; c < 18; c++) {
      target71[yearKey][c] = r26[yearKey][c];
    }
  }
}

console.log("=== STEP 3: Compiling index.html with Tone & Manner ===");

const regionMeta = {
  'EU': { en: 'Europe HQ (유럽 본부)', kr: 'EUROPE/CIS' },
  'LGEAG': { en: 'Austria (오스트리아)', kr: '오스트리아 지점' },
  'LGEBN': { en: 'Benelux (베네룩스)', kr: '베네룩스 지점' },
  'LGECK': { en: 'Czech (체코)', kr: '체코 법인' },
  'LGEDG': { en: 'Germany (독일)', kr: '독일 법인' },
  'Swiss': { en: 'Switzerland (스위스)', kr: '스위스 지점' },
  'LGEES': { en: 'Spain (스페인)', kr: '스페인 법인' },
  'LGEFS': { en: 'France (프랑스)', kr: '프랑스 법인' },
  'LGEHS': { en: 'Greece (그리스)', kr: '그리스 법인' },
  'LGEIS': { en: 'Italy (이탈리아)', kr: '이탈리아 법인' },
  'LGELA': { en: 'Baltics (발트 3국)', kr: '발트 지점' },
  'LGEMK': { en: 'Hungary (헝가리)', kr: '헝가리 법인' },
  'LGEPL': { en: 'Poland (폴란드)', kr: '폴란드 법인' },
  'LGEPT': { en: 'Portugal (포르투갈)', kr: '포르투갈 지점' },
  'LGERO': { en: 'Romania (루마니아)', kr: '루마니아 지점' },
  'LGESW': { en: 'Sweden (스웨덴)', kr: '스웨덴 법인' },
  'LGEUK': { en: 'United Kingdom (영국)', kr: '영국 법인' },
  'CIS': { en: 'CIS HQ (러시아/CIS 본부)', kr: '유라시아 전체 (HQ)' },
  'LGEAK': { en: 'Kazakhstan (알마티)', kr: '알마티 법인' },
  'LGERA': { en: 'Russia (러시아)', kr: '러시아 법인' },
  'LGEUR': { en: 'Ukraine (우크라이나)', kr: '우크라이나 지점' }
};

const htmlTemplate = `<!DOCTYPE html>
<html class="light" lang="ko">
<head>
    <meta charset="utf-8"/>
    <meta content="width=device-width, initial-scale=1.0" name="viewport"/>
    <title>Executive Portal | LGE Europe TV KPI Dashboard</title>
    <script src="https://cdn.tailwindcss.com?plugins=forms,container-queries"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;0,8..60,700;1,8..60,400&amp;family=Noto+Sans+KR:wght@300;400;500;700&amp;display=swap" rel="stylesheet"/>
    <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&amp;display=swap" rel="stylesheet"/>
    <script id="tailwind-config">
        tailwind.config = {
            darkMode: "class",
            theme: {
                extend: {
                    "colors": {
                        "primary": "#051c2c",
                        "on-primary": "#ffffff",
                        "secondary": "#a50034",
                        "surface": "#f6faff",
                        "surface-variant": "#ebf5ff",
                        "outline": "#73777d",
                        "outline-variant": "#c3c7cc",
                        "error": "#e11d48",
                        "background": "#ffffff",
                        "primary-hover": "#860027",
                        "teal-accent": "#00a3a3",
                        "crimson-accent": "#e11d48"
                    },
                    "borderRadius": {
                        "DEFAULT": "0.125rem",
                        "lg": "0.25rem",
                        "xl": "0.5rem"
                    },
                    "fontFamily": {
                        "headline": ["\\"Source Serif 4\\"", "serif"],
                        "body": ["Inter", "Noto Sans KR", "sans-serif"]
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Inter', 'Noto Sans KR', sans-serif; background-color: #f6faff; color: #051c2c; }
        h1, h2, h3, h4, h5, h6 { font-family: 'Source Serif 4', serif; }
        .sidebar-item.active { background-color: rgba(255, 255, 255, 0.1); border-right: 4px solid #ffffff; }
        .country-node.active > button { background-color: rgba(255, 255, 255, 0.05); color: #ffffff; }
        html, body { overflow-x: hidden; width: 100%; }
        .sticky-col { position: sticky; left: 0; background-color: #ffffff; z-index: 10; }
        .sticky-col-header { position: sticky; left: 0; z-index: 20; }
        .table-container { position: relative; overflow-x: auto; }
        th { position: sticky; top: 0; background-color: #051c2c; color: #ffffff; z-index: 15; }
    </style>
</head>
<body class="flex min-h-screen">
    <!-- Sidebar Navigation -->
    <aside class="w-72 bg-primary text-white fixed h-screen flex flex-col z-50 shadow-lg">
        <div class="p-8 border-b border-white/10 flex-shrink-0">
            <div class="flex items-center gap-2 mb-2">
                <span class="material-symbols-outlined text-white/80">insights</span>
                <h1 class="text-xl font-bold tracking-tight">TV KPI Portal</h1>
            </div>
            <p class="text-[10px] text-white/50 uppercase tracking-[0.2em] font-medium">TV EUROPE/CIS TV PORTAL</p>
        </div>
        
        <nav class="flex-1 py-6 space-y-4 overflow-y-auto" id="sidebar-nav">
            <!-- Europe HQ -->
            <div>
                <div class="px-8 py-2 text-[10px] font-bold text-white/40 uppercase tracking-widest mb-1">Europe HQ (유럽 본부)</div>
                <div class="space-y-0.5" id="group-europe"></div>
            </div>
            <!-- CIS HQ -->
            <div>
                <div class="px-8 py-2 text-[10px] font-bold text-white/40 uppercase tracking-widest mb-1">CIS HQ (러시아/CIS)</div>
                <div class="space-y-0.5" id="group-cis"></div>
            </div>
        </nav>
        
        <div class="p-8 border-t border-white/10 bg-black/10 flex flex-col gap-1 text-[10px] text-white/40 font-sans tracking-tight leading-relaxed flex-shrink-0">
            <p class="font-bold text-white/50 text-[11px] mb-1">System Admin</p>
            <p class="text-white/60">Harry Park</p>
            <p><a href="mailto:harry.park@lge.com" class="hover:text-white text-white/50 underline decoration-white/20 transition-all">harry.park@lge.com</a></p>
        </div>
    </aside>

    <!-- Main Content Area -->
    <main class="ml-72 flex-1 flex flex-col min-h-screen min-w-0">
        <!-- Top Bar Header -->
        <header class="h-20 bg-white border-b border-slate-200 flex items-center justify-between px-12 sticky top-0 z-40">
            <div class="flex items-center gap-6">
                <span class="text-xs font-bold tracking-widest text-slate-400 uppercase">EUROPE/CIS TV PORTAL</span>
                <div class="h-4 w-px bg-slate-200"></div>
                <p class="text-sm font-semibold text-primary" id="page-indicator">로딩 중...</p>
            </div>
            <div class="flex items-center gap-6">
                <!-- Year Selector Tabs -->
                <div class="flex bg-slate-100 p-1 rounded">
                    <button class="px-4 py-1.5 text-xs font-bold rounded transition-all active-year-tab bg-primary text-white" id="tab-2026" onclick="setYear(2026)">2026년 실적</button>
                    <button class="px-4 py-1.5 text-xs font-bold rounded transition-all text-slate-600 hover:text-primary" id="tab-2025" onclick="setYear(2025)">2025년 실적</button>
                </div>
                <div class="h-8 w-px bg-slate-200"></div>
                <div class="text-right">
                    <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Dashboard Mode</p>
                    <p class="text-sm font-bold text-secondary">Interactive Web</p>
                </div>
            </div>
        </header>

        <!-- View Container -->
        <div class="p-12 w-full flex-1 flex flex-col gap-10">
            <!-- Welcome Banner -->
            <div class="bg-primary text-white p-8 rounded-lg shadow-md relative overflow-hidden flex-shrink-0">
                <div class="relative z-10">
                    <h2 class="text-3xl font-headline font-bold mb-3" id="banner-title">유럽/CIS TV Biz. KPI Monitoring Dashboard</h2>
                    <p class="text-xs text-white/70 leading-relaxed max-w-4xl" id="banner-desc">유럽 및 CIS 주요 법인/지점의 TV 사업 핵심 실적 지표(수량, 재고, 손익, Market Share)를 모니터링하는 인터랙티브 대시보드입니다. 좌측 법인/지점 목록에서 법인/지점을 선택하면 Excel 수식 엔진에 의해 산출된 실적 및 YoY 전년비 분석 결과가 실시간으로 로드됩니다.</p>
                </div>
                <div class="absolute -right-24 -bottom-24 w-80 h-80 bg-white/5 rounded-full blur-2xl pointer-events-none"></div>
            </div>

            <!-- KPI Cards Section -->
            <div class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-6 gap-4 flex-shrink-0" id="kpi-cards-grid">
                <!-- Cards injected here -->
            </div>

            <!-- Table Selection Menu (Tabs) -->
            <div class="flex border-b border-slate-200 gap-2 flex-shrink-0">
                <button id="tab-btn-psi" class="px-6 py-2.5 text-sm font-bold border-b-2 border-secondary text-primary transition-all" onclick="switchTableTab('psi')">PSI 지표 (PSI Index)</button>
                <button id="tab-btn-pnl" class="px-6 py-2.5 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-primary transition-all" onclick="switchTableTab('pnl')">손익 지표 (P&L Index)</button>
                <button id="tab-btn-comp" class="px-6 py-2.5 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-primary transition-all" onclick="switchTableTab('comp')">경쟁 지표 (Competitiveness Indicator)</button>
            </div>

            <!-- Data Tables Section -->
            <div class="flex flex-col gap-8">
                <!-- Table 1: PSI -->
                <div id="section-psi" class="bg-white border border-slate-200 rounded-lg shadow-sm p-6">
                    <div class="flex justify-between items-center mb-4">
                        <div class="flex items-center gap-3">
                            <span class="h-5 w-1 bg-primary"></span>
                            <h4 class="text-lg font-headline font-bold text-primary">PSI 지표 (PSI Index)</h4>
                        </div>
                        <span class="text-xs text-slate-400 font-medium">단위: 수량(대/ea), WOS(주/weeks)</span>
                    </div>
                    <div class="table-container border border-slate-100 rounded">
                        <table class="min-w-full border-collapse text-left text-xs" id="table-psi">
                            <!-- Injected by JS -->
                        </table>
                    </div>
                </div>

                <!-- Table 2: P&L -->
                <div id="section-pnl" class="bg-white border border-slate-200 rounded-lg shadow-sm p-6 hidden">
                    <div class="flex justify-between items-center mb-4">
                        <div class="flex items-center gap-3">
                            <span class="h-5 w-1 bg-secondary"></span>
                            <h4 class="text-lg font-headline font-bold text-primary">손익 지표 (P&L Index)</h4>
                        </div>
                        <span class="text-xs text-slate-400 font-medium">단위: 금액($, K$), 비율(%)</span>
                    </div>
                    <div class="table-container border border-slate-100 rounded">
                        <table class="min-w-full border-collapse text-left text-xs" id="table-pnl">
                            <!-- Injected by JS -->
                        </table>
                    </div>
                </div>

                <!-- Table 3: Competition -->
                <div id="section-comp" class="bg-white border border-slate-200 rounded-lg shadow-sm p-6 hidden">
                    <div class="flex justify-between items-center mb-4">
                        <div class="flex items-center gap-3">
                            <span class="h-5 w-1 bg-teal-600"></span>
                            <h4 class="text-lg font-headline font-bold text-primary">경쟁 지표 (Competitiveness Indicator)</h4>
                        </div>
                        <span class="text-xs text-slate-400 font-medium">단위: 비율(%), 수량(대/ea), 지수(index)</span>
                    </div>
                    <div class="table-container border border-slate-100 rounded">
                        <table class="min-w-full border-collapse text-left text-xs" id="table-comp">
                            <!-- Injected by JS -->
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <!-- Embed JSON Data -->
    <script>
        const regionMeta = ${JSON.stringify(regionMeta)};
        const kpiData = ${JSON.stringify(processedData)};
        
        let currentRegion = 'Swiss';
        let currentYear = 2026;
        let currentTableTab = 'psi';
        
        window.switchTableTab = function(tabName) {
            currentTableTab = tabName;
            
            // Toggle sections
            document.getElementById('section-psi').classList.toggle('hidden', tabName !== 'psi');
            document.getElementById('section-pnl').classList.toggle('hidden', tabName !== 'pnl');
            document.getElementById('section-comp').classList.toggle('hidden', tabName !== 'comp');
            
            // Toggle button styles
            const tabs = ['psi', 'pnl', 'comp'];
            tabs.forEach(t => {
                const btn = document.getElementById('tab-btn-' + t);
                if (btn) {
                    if (t === tabName) {
                        btn.className = "px-6 py-2.5 text-sm font-bold border-b-2 border-secondary text-primary transition-all";
                    } else {
                        btn.className = "px-6 py-2.5 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-primary transition-all";
                    }
                }
            });
        };
        
        const colHeaders = [
            "1월", "2월", "3월", "4월", "5월", "6월", "상반기",
            "7월", "8월", "9월", "10월", "11월", "12월",
            "1Q", "2Q", "3Q", "4Q", "TTL"
        ];
        
        // Render Sidebar
        function renderSidebar() {
            const europeGroup = document.getElementById('group-europe');
            const cisGroup = document.getElementById('group-cis');
            europeGroup.innerHTML = '';
            cisGroup.innerHTML = '';
            
            const listEurope = ['EU', 'LGEAG', 'LGEBN', 'LGECK', 'LGEDG', 'Swiss', 'LGEES', 'LGEFS', 'LGEHS', 'LGEIS', 'LGELA', 'LGEMK', 'LGEPL', 'LGEPT', 'LGERO', 'LGESW', 'LGEUK'];
            const listCis = ['CIS', 'LGEAK', 'LGERA', 'LGEUR'];
            
            listEurope.forEach(code => {
                europeGroup.appendChild(createSidebarButton(code));
            });
            listCis.forEach(code => {
                cisGroup.appendChild(createSidebarButton(code));
            });
        }
        
        function createSidebarButton(code) {
            const meta = regionMeta[code];
            const btn = document.createElement('button');
            btn.className = "w-full flex items-center justify-between px-6 py-2.5 text-left text-white/70 hover:text-white hover:bg-white/5 transition-all text-xs font-medium sidebar-btn-" + code;
            if (code === currentRegion) {
                btn.className += " active bg-white/10 text-white font-bold border-r-4 border-r-white";
            }
            
            const shortCode = code.replace('LGE', '');
            btn.innerHTML = \`
                <span class="flex items-center gap-2.5">
                    <span class="text-[9px] font-mono font-bold bg-white/10 text-white/80 px-1.5 py-0.5 rounded border border-white/10 uppercase tracking-wider">\${shortCode}</span>
                    <span>\${meta.en.split(' (')[0]}</span>
                </span>
            \`;
            btn.onclick = () => selectRegion(code);
            return btn;
        }
        
        function selectRegion(code) {
            currentRegion = code;
            document.querySelectorAll('#sidebar-nav button').forEach(btn => {
                btn.classList.remove('active', 'bg-white/10', 'font-bold', 'border-r-4', 'border-r-white');
                btn.classList.add('text-white/70');
            });
            const activeBtn = document.querySelector('.sidebar-btn-' + code);
            if (activeBtn) {
                activeBtn.classList.add('active', 'bg-white/10', 'font-bold', 'border-r-4', 'border-r-white');
                activeBtn.classList.remove('text-white/70');
            }
            updateView();
        }
        
        function setYear(year) {
            currentYear = year;
            document.getElementById('tab-2026').className = year === 2026 ? "px-4 py-1.5 text-xs font-bold rounded transition-all bg-primary text-white" : "px-4 py-1.5 text-xs font-bold rounded transition-all text-slate-600 hover:text-primary";
            document.getElementById('tab-2025').className = year === 2025 ? "px-4 py-1.5 text-xs font-bold rounded transition-all bg-primary text-white" : "px-4 py-1.5 text-xs font-bold rounded transition-all text-slate-600 hover:text-primary";
            
            updateView();
        }
        
        function formatVal(val, type, label = '', rowIndex = null) {
            if (val === null || val === undefined || isNaN(val)) return '-';
            
            const isPnlRow = rowIndex !== null && rowIndex >= 38 && rowIndex <= 67;
            
            if (isPnlRow) {
                if (rowIndex >= 41 && rowIndex <= 43) {
                    const workingVal = val / 1e6;
                    return '\\$' + workingVal.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
                }
                if (rowIndex >= 53 && rowIndex <= 55) {
                    const workingVal = val / 1e6;
                    return '\\$' + workingVal.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
                }
                const isPnlPercent = (rowIndex >= 44 && rowIndex <= 46) || (rowIndex >= 62 && rowIndex <= 67);
                if (isPnlPercent) {
                    return (val * 100).toFixed(1) + '%';
                }
                const valInM = val / 1e6;
                return '\\$' + valInM.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'M';
            }
            
            const isWos = type === 'w' || rowIndex === 32 || rowIndex === 33;
            const isAsp = label.toLowerCase().includes('asp');
            let workingVal = val;
            if (isAsp) {
                workingVal = val / 1e6;
            }
            
            if (type === 'p') {
                const decimalPlaces = (rowIndex === 69 || rowIndex === 71) ? 0 : 1;
                let pctVal = workingVal * 100;
                const threshold = decimalPlaces === 0 ? 0.5 : 0.05;
                if (Math.abs(pctVal) < threshold) {
                    pctVal = 0;
                }
                return pctVal.toFixed(decimalPlaces) + '%';
            }
            if (type === 'm') {
                if (isAsp) {
                    return '\\$' + workingVal.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
                }
                if (Math.abs(workingVal) >= 1e6) {
                    const valInM = workingVal / 1e6;
                    return '\\$' + valInM.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'M';
                }
                return '\\$' + workingVal.toLocaleString(undefined, {maximumFractionDigits: 0});
            }
            if (type === 'i' || isWos) {
                return workingVal.toFixed(1);
            }
            return workingVal.toLocaleString(undefined, {maximumFractionDigits: 0});
        }

        function formatYoY(val, type) {
            if (val === null || val === undefined || isNaN(val)) return '-';
            const sign = val > 0 ? '▲ ' : (val < 0 ? '▼ ' : '');
            const absVal = Math.abs(val);
            if (type === 'p') {
                return sign + absVal.toFixed(1) + '%p';
            }
            return sign + absVal.toFixed(1) + '%';
        }

        function formatDiffVal(val, type, label = '', rowIndex = null) {
            if (val === null || val === undefined || isNaN(val)) return '-';
            const sign = val > 0 ? '+' : '';
            const formatted = formatVal(val, type, label, rowIndex);
            if (formatted === '-') return '-';
            if (formatted.startsWith('-') || formatted.startsWith('-$') || formatted.startsWith('▼') || formatted.startsWith('▲')) {
                return formatted;
            }
            return sign + formatted;
        }

        window.toggleSubRows = function(rowIndex) {
            const subRows = document.getElementsByClassName('sub-row-' + rowIndex);
            const icon = document.getElementById('icon-' + rowIndex);
            Array.from(subRows).forEach(row => {
                row.classList.toggle('hidden');
            });
            if (icon) {
                if (icon.innerText === '▶') {
                    icon.innerText = '▼';
                } else {
                    icon.innerText = '▶';
                }
            }
        };
        
        function getYoYColorClass(val, type) {
            if (val === null || val === undefined || isNaN(val) || val === 0) return '';
            // For WOS, inventory, lower is better
            const isReverse = type === 'w' || type === 'i';
            const isPositive = isReverse ? val < 0 : val > 0;
            return isPositive ? 'bg-green-50 text-green-700 font-medium' : 'bg-red-50 text-red-700 font-medium';
        }

        function getRowDisplayNameAndIndent(labels) {
            let indent = 0;
            let text = '';
            for (let i = 4; i >= 0; i--) {
                if (labels[i] !== null && labels[i] !== undefined && String(labels[i]).trim() !== '') {
                    indent = i;
                    text = String(labels[i]).trim();
                    break;
                }
            }
            // Align "시장판가 (시장)" and "M/S (LG)" with "전시지수 (TTL)" (indent = 2)
            if (text === "시장판가 (시장)" || text === "M/S (LG)") {
                indent = 2;
            }
            return { text, indent };
        }
        
        function updateView() {
            const meta = regionMeta[currentRegion];
            document.getElementById('page-indicator').innerText = \`\${meta.kr} TV KPI Dashboard\`;
            document.getElementById('banner-title').innerText = \`\${meta.kr} TV Biz. KPI Monitoring Dashboard\`;
            
            const rows = kpiData[currentRegion] || [];
            
            // 1. Render Summary Cards
            renderSummaryCards(rows);
            
            // 2. Render Tables
            // Filter and sort PSI rows: we want rows 19 to 33 (excluding 26) and rows 69, 71.
            // But we must sort them so that 69 and 71 appear right before 28!
            const psiRows = rows.filter(r => 
                ((r.index >= 19 && r.index <= 33) && r.index !== 26) ||
                (r.index === 69 || r.index === 71)
            );
            psiRows.sort((a, b) => {
                const getSortOrder = (idx) => {
                    if (idx === 69) return 27.1;
                    if (idx === 71) return 27.2;
                    return idx;
                };
                return getSortOrder(a.index) - getSortOrder(b.index);
            });
            const pnlRows = rows.filter(r => r.index >= 38 && r.index <= 67);
            const compRows = rows.filter(r => r.index >= 73 && r.index <= 93);
            
            renderTable('table-psi', psiRows);
            renderTable('table-pnl', pnlRows);
            renderTable('table-comp', compRows);
        }

        function renderSummaryCards(rows) {
            const grid = document.getElementById('kpi-cards-grid');
            grid.innerHTML = '';
            
            // Key row indices in sheet (Net Sales -> 영업이익 -> 영업이익율 -> Sell-in -> Sell-out -> 유통재고)
            const cardRowDefs = [
                { title: 'Net Sales', index: 50, unit: 'K\$' },
                { title: '영업이익', index: 59, unit: 'K\$' },
                { title: '영업이익율', index: 65, unit: '%' },
                { title: 'Sell-in 수량', index: 19, unit: '대' },
                { title: 'Sell-out 수량', index: 24, unit: '대' },
                { title: '유통재고', index: 28, unit: '대' }
            ];
            
            cardRowDefs.forEach(def => {
                const row = rows.find(r => r.index === def.index);
                if (!row) return;
                
                const valArray = currentYear === 2026 ? row.y26 : row.y25;
                const yoyArray = currentYear === 2026 ? row.y26_yoy : row.y25_yoy;
                
                // YTD (June) logic:
                // Flow metrics (Sell-in, Sell-out, Net Sales, 영업이익, 영업이익율) use H1 (index 6, Column R)
                // Stock/Snapshot metrics (유통재고) use June (index 5, Column Q)
                const isStockMetric = def.index === 28;
                const targetIdx = isStockMetric ? 5 : 6;
                const periodLabel = isStockMetric ? 'YTD (6월 말)' : 'YTD (6월 누적)';
                
                let val = valArray ? valArray[targetIdx] : null;
                let yoy = yoyArray ? yoyArray[targetIdx] : null;
                
                let valStr = formatVal(val, row.type);
                const yoyStr = formatYoY(yoy, row.type);
                const yoyColor = getYoYColorClass(yoy, row.type);
                const borderAccent = (def.index === 59 || def.index === 65) ? 'border-l-4 border-l-secondary' : 'border-l-4 border-l-teal-accent';
                
                // Calculate absolute change compared to previous year
                let diffStr = '';
                if (val !== null && val !== undefined) {
                    const prevValArray = currentYear === 2026 ? row.y25 : row.y24;
                    const prevVal = prevValArray ? prevValArray[targetIdx] : null;
                    
                    if (prevVal !== null && prevVal !== undefined) {
                        const diff = val - prevVal;
                        const sign = diff >= 0 ? '+' : '-';
                        const absDiff = Math.abs(diff);
                        
                        if (def.index === 50 || def.index === 59) {
                            // Currency: Net Sales & Operating Income
                            const absM = absDiff / 1e6;
                            diffStr = sign + '\\$' + absM.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'M';
                        } else if (def.index === 19 || def.index === 24 || def.index === 28) {
                            // Quantity: Sell-in, Sell-out, 유통재고
                            diffStr = sign + absDiff.toLocaleString(undefined, { maximumFractionDigits: 0 });
                        }
                    }
                }
                
                let badgeText = yoyStr;
                if (diffStr && yoy !== null) {
                    badgeText = diffStr + ' (' + yoyStr + ')';
                }
                
                const card = document.createElement('div');
                card.className = \`bg-white p-5 rounded border border-slate-200 hover:shadow-md transition-all flex flex-col justify-between \${borderAccent}\`;
                card.innerHTML = \`
                    <div class="flex flex-col gap-1.5 items-start">
                        <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">\${def.title}</span>
                        <span class="text-2xl font-bold text-primary font-headline tracking-tight">\${valStr}</span>
                        <span class="text-[10px] text-slate-500 font-semibold bg-slate-50 px-2 py-0.5 rounded">\${periodLabel}</span>
                        <span class="text-[10px] px-2 py-0.5 rounded border border-slate-100 mt-0.5 \${yoyColor}">\${badgeText}</span>
                    </div>
                \`;
                grid.appendChild(card);
            });
        }

                function renderTable(tableId, rows) {
            const table = document.getElementById(tableId);
            if (!table) return;
            table.innerHTML = '';
            
            // Header
            let html = '<thead><tr>';
            html += '<th class="sticky-col-header text-left px-4 py-3 bg-primary text-white font-bold" style="min-width: 260px;">지표 (Indicator)</th>';
            colHeaders.forEach(h => {
                const bgClass = h === '상반기' || h === 'TTL' || h.includes('Q') ? 'bg-slate-800' : 'bg-primary';
                html += \`<th class="text-center px-2 py-3 \${bgClass} text-white font-bold">\${h}</th>\`;
            });
            html += '</tr></thead><tbody>';
            
            // Rows
            rows.forEach((row, idx) => {
                const { text, indent } = getRowDisplayNameAndIndent(row.labels);
                if (!text) return; // Skip separator/empty rows
                
                const valArray = currentYear === 2026 ? row.y26 : row.y25;
                const prevValArray = currentYear === 2026 ? row.y25 : row.y24;
                
                const indentClass = \`pl-\${indent * 4 + 4}\`;
                const textStyle = indent === 0 ? 'font-bold text-primary text-xs' : (indent === 1 ? 'font-medium text-slate-700 text-[11px]' : 'text-slate-600 text-[11px]');
                const trBg = idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50';
                
                // Clickable main row
                html += \`<tr class="\${trBg} hover:bg-slate-100/70 border-b border-slate-100 cursor-pointer" onclick="toggleSubRows('\${row.index}')">\`;
                html += \`<td class="sticky-col px-4 py-2 border-r border-slate-100 \${indentClass} \${textStyle}" style="min-width: 260px; left:0;">\`;
                html += \`<span class="mr-1.5 inline-block text-[9px] text-slate-400 select-none transition-transform" id="icon-\${row.index}">▶</span>\`;
                html += \`\${text}</td>\`;
                
                for (let c = 0; c < 18; c++) {
                    const val = valArray ? valArray[c] : null;
                    const valStr = formatVal(val, row.type, text, row.index);
                    const tdBg = colHeaders[c] === '상반기' || colHeaders[c] === 'TTL' || colHeaders[c].includes('Q') ? 'bg-slate-50 font-bold' : '';
                    html += \`<td class="text-right px-2 py-2 border-r border-slate-100 \${tdBg} \${textStyle}">\${valStr}</td>\`;
                }
                html += '</tr>';
                
                // Collapsible row 1: Last Year (전년)
                const subRowClass = \`sub-row-\${row.index}\`;
                const childIndentClass = \`pl-\${indent * 4 + 12}\`;
                const childTextStyle = 'text-slate-400 font-normal italic text-[10px] bg-slate-50/30';
                
                html += \`<tr class="\${subRowClass} hidden bg-slate-50/20 border-b border-slate-50">\`;
                html += \`<td class="sticky-col px-4 py-1.5 border-r border-slate-100 \${childIndentClass} \${childTextStyle}" style="min-width: 260px; left:0;">└ 전년 (Prev Year)</td>\`;
                for (let c = 0; c < 18; c++) {
                    const val = prevValArray ? prevValArray[c] : null;
                    const valStr = formatVal(val, row.type, text, row.index);
                    const tdBg = colHeaders[c] === '상반기' || colHeaders[c] === 'TTL' || colHeaders[c].includes('Q') ? 'bg-slate-50/30 font-semibold' : '';
                    html += \`<td class="text-right px-2 py-1.5 border-r border-slate-100 \${tdBg} \${childTextStyle}">\${valStr}</td>\`;
                }
                html += '</tr>';
                
                // Collapsible row 2: YoY Diff (전년비 수량차/금액차)
                html += \`<tr class="\${subRowClass} hidden bg-slate-50/20 border-b border-slate-100">\`;
                html += \`<td class="sticky-col px-4 py-1.5 border-r border-slate-100 \${childIndentClass} \${childTextStyle}" style="min-width: 260px; left:0;">└ 전년비 (YoY Diff)</td>\`;
                for (let c = 0; c < 18; c++) {
                    const valCur = valArray ? valArray[c] : null;
                    const valPrev = prevValArray ? prevValArray[c] : null;
                    const diff = (valCur !== null && valPrev !== null) ? (valCur - valPrev) : null;
                    const valStr = formatDiffVal(diff, row.type, text, row.index);
                    const tdBg = colHeaders[c] === '상반기' || colHeaders[c] === 'TTL' || colHeaders[c].includes('Q') ? 'bg-slate-50/30 font-semibold' : '';
                    html += \`<td class="text-right px-2 py-1.5 border-r border-slate-100 \${tdBg} \${childTextStyle}">\${valStr}</td>\`;
                }
                html += '</tr>';
            });
            
            html += '</tbody>';
            table.innerHTML = html;
        }

        // Init
        renderSidebar();
        updateView();
        </script>
</body>
</html>`;

fs.writeFileSync(finalHtmlPath, htmlTemplate, 'utf8');
console.log("=== Build complete! generated: index.html ===");
