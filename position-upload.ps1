# position-upload.ps1 — 회사 PC용 단일 파일: 포지션 엑셀 추출 → FICC 모니터 API로 업로드.
#  파일 반출 없이 회사 PC에서 실행 (K-Bond 수집기와 같은 아웃바운드 HTTP 패턴).
#  사용: 1) 아래 설정 3줄 확인  2) powershell -File position-upload.ps1  3) 작업스케줄러 등록(매일 장마감 후)
#  ※ 이 파일은 UTF-8 BOM 인코딩 유지 필수 (PS 5.1 한글).
$ErrorActionPreference = "Continue"

# ── 설정 (회사 PC 환경에 맞게) ───────────────────────────────────────────────
$ENDPOINT = "https://bondmonitoring.onrender.com/api/import/positions"
$TOKEN    = "CHANGE_ME"                                  # ★ FICC 모니터 IMPORT_TOKEN 으로 교체 (여기만 수정)
$FILE_DIR = "C:\kbond-collector\positions"               # App.금리차익.포지션.*.xlsm 을 넣어두는 폴더
$BOOK     = "B020105"                                    # 금리차익 북

# ── 최신 포지션 파일 탐색 ────────────────────────────────────────────────────
$f = Get-ChildItem $FILE_DIR -File -Filter "App.금리차익.포지션.*.xlsm" -Recurse -Depth 2 -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $f) { Write-Host "[upload] 포지션 파일 없음: $FILE_DIR"; exit 2 }
Write-Host ("[upload] file: " + $f.FullName)

$xl = $null; $opened = $false
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $opened = $true }
$wb = $null; foreach ($b in $xl.Workbooks) { if ($b.FullName -eq $f.FullName) { $wb = $b } }
$wbOpened = $false
if (-not $wb) { $wb = $xl.Workbooks.Open($f.FullName, 0, $true); $wbOpened = $true }

function Sheet($wb, $name) { foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { return $s } }; return $null }
function OaYmd($v) { if ($v -is [double] -and $v -gt 20000) { return [DateTime]::FromOADate($v).ToString("yyyy-MM-dd") }; return "" }
function Num($v) { if ($v -is [double] -or $v -is [int]) { return [string]$v }; return "" }

# ── 1) 금리민감도 ────────────────────────────────────────────────────────────
$ms = Sheet $wb "금리민감도"
$asOf = OaYmd $ms.Cells.Item(1,1).Value2
$nR = $ms.UsedRange.Rows.Count
$arr = $ms.Range($ms.Cells(1,1), $ms.Cells($nR, 38)).Value2
$riskLines = New-Object System.Collections.Generic.List[string]
$riskLines.Add("asOf,sht,bookCode,fundCode,symbol,name,riskFactor,kind,assetClass,strategy,ytm,cpn,modDur,price,maturity,bondClass,rating,quantity,bookValue,carry,pv01,t1d,t3m,t6m,t9m,t1y,t18m,t2y,t30m,t3y,t4y,t5y,t7y,t10y,t12y,t15y,t20y,t30y")
for ($r = 2; $r -le $nR; $r++) {
  if ([string]$arr[$r,2] -ne $BOOK) { continue }
  $name = ([string]$arr[$r,5]) -replace '[",]', ' '
  $vals = @($asOf, [string]$arr[$r,1], [string]$arr[$r,2], [string]$arr[$r,3], [string]$arr[$r,4], $name,
    [string]$arr[$r,6], [string]$arr[$r,7], [string]$arr[$r,8], [string]$arr[$r,9],
    (Num $arr[$r,10]), (Num $arr[$r,11]), (Num $arr[$r,12]), (Num $arr[$r,13]), (OaYmd $arr[$r,14]),
    ([string]$arr[$r,15] -replace ',',' '), ([string]$arr[$r,16] -replace ',',' '),
    (Num $arr[$r,17]), (Num $arr[$r,19]), (Num $arr[$r,20]), (Num $arr[$r,21]))
  for ($c = 22; $c -le 38; $c++) { $vals += (Num $arr[$r,$c]) }
  $riskLines.Add(($vals -join ","))
}
Write-Host ("[upload] 금리민감도 " + ($riskLines.Count - 1) + "행 (asOf " + $asOf + ")")

# ── 2) 채무증권 ──────────────────────────────────────────────────────────────
$bs = Sheet $wb "채무증권"
$nR = $bs.UsedRange.Rows.Count
$arr = $bs.Range($bs.Cells(1,1), $bs.Cells($nR, 26)).Value2
$bondLines = New-Object System.Collections.Generic.List[string]
$bondLines.Add("asOf,sht,fundCode,symbol,name,kind,riskFactor,rating,rateType,buyDate,maturity,quantity,evalPrice,modDur,ytm,carry,buyYield,cpn,cpnFreq,buyPrice,bookValue,evalValue")
for ($r = 2; $r -le $nR; $r++) {
  if ([string]$arr[$r,2] -ne $BOOK) { continue }
  $name = ([string]$arr[$r,5]) -replace '[",]', ' '
  $bondLines.Add((@($asOf, [string]$arr[$r,1], [string]$arr[$r,3], [string]$arr[$r,4], $name, [string]$arr[$r,6], [string]$arr[$r,7],
    ([string]$arr[$r,8] -replace ',',' '), [string]$arr[$r,9], (OaYmd $arr[$r,10]), (OaYmd $arr[$r,11]),
    (Num $arr[$r,13]), (Num $arr[$r,16]), (Num $arr[$r,17]), (Num $arr[$r,18]), (Num $arr[$r,19]), (Num $arr[$r,20]),
    (Num $arr[$r,21]), (Num $arr[$r,22]), (Num $arr[$r,24]), (Num $arr[$r,25]), (Num $arr[$r,26])) -join ","))
}
Write-Host ("[upload] 채무증권 " + ($bondLines.Count - 1) + "행")

# ── 3) 캐리 시나리오 요약 ────────────────────────────────────────────────────
$cs = Sheet $wb "캐리 시나리오"
$arr = $cs.Range($cs.Cells(1,1), $cs.Cells(30, 30)).Value2
$carryLines = New-Object System.Collections.Generic.List[string]
$carryLines.Add("asOf,key,value,label")
function AddMeta($lines, $asOf, $key, $v, $label) { if ($v -is [double] -or $v -is [int]) { $lines.Add("$asOf,$key,$v," + ($label -replace ',',' ')) } }
$posCols = @("total","govt","credit","bankCap","cpStn","abs")
for ($r = 2; $r -le 7; $r++) {
  $strat = [string]$arr[$r,1]; if (-not $strat) { continue }
  for ($c = 2; $c -le 7; $c++) { AddMeta $carryLines $asOf ("pos." + $strat + "." + $posCols[$c-2]) $arr[$r,$c] ($strat + " " + $posCols[$c-2]) }
}
for ($c = 2; $c -le 7; $c++) { AddMeta $carryLines $asOf ("limit." + $posCols[$c-2]) $arr[9,$c] ("한도 " + $posCols[$c-2]) }
$ladderCols = @("under1y","y1","y2","y3","y5","y10","y30","sum")
for ($r = 14; $r -le 20; $r++) {
  $cls = [string]$arr[$r,1]; if (-not $cls) { continue }
  $cls = $cls -replace '[,/]', '_'
  for ($c = 2; $c -le 9; $c++) { AddMeta $carryLines $asOf ("ladder." + $cls + "." + $ladderCols[$c-2]) $arr[$r,$c] ($cls + " " + $ladderCols[$c-2]) }
}
for ($r = 2; $r -le 7; $r++) {
  $ten = $arr[$r,12]; if (-not ($ten -is [double])) { continue }
  AddMeta $carryLines $asOf ("ytmcost." + $ten + ".ktb") $arr[$r,13] ("비용대비YTM " + $ten + "Y KTB")
  AddMeta $carryLines $asOf ("ytmcost." + $ten + ".irs") $arr[$r,14] ("IRS")
  AddMeta $carryLines $asOf ("ytmcost." + $ten + ".bs") $arr[$r,15] ("BS")
}
$carryRows = @(@(10,"base"),@(11,"base"),@(12,"base"),@(13,"base"),@(17,"roll"),@(18,"roll"),@(19,"roll"),@(20,"roll"))
foreach ($cr in $carryRows) {
  $r = $cr[0]; $mode = $cr[1]
  $g = ([string]$arr[$r,11]) -replace '[()\s]', ''; $a = ([string]$arr[$r,13])
  if (-not $g -and -not $a) { continue }
  $kg = if ($g) { $g } else { "x" }
  AddMeta $carryLines $asOf ("carry." + $mode + "." + $kg + ".theta") $arr[$r,12] ($mode + " " + $kg + " 부채theta")
  AddMeta $carryLines $asOf ("carry." + $mode + "." + $kg + "." + $a + ".assetCarry") $arr[$r,14] ($mode + " " + $kg + " " + $a + " 자산carry")
  AddMeta $carryLines $asOf ("carry." + $mode + "." + $kg + "." + $a + ".funding") $arr[$r,15] ($mode + " 조달비용")
  AddMeta $carryLines $asOf ("carry." + $mode + "." + $kg + ".total") $arr[$r,16] ($mode + " Total")
}
Write-Host ("[upload] 캐리요약 " + ($carryLines.Count - 1) + "값")

# ── 4) 펀드→전략 ─────────────────────────────────────────────────────────────
$fc = Sheet $wb "Query.펀드코드.공학팀"
$nR = $fc.UsedRange.Rows.Count
$arr = $fc.Range($fc.Cells(1,1), $fc.Cells($nR, 3)).Value2
$fundLines = New-Object System.Collections.Generic.List[string]
$fundLines.Add("bookCode,fundCode,strategy")
for ($r = 2; $r -le $nR; $r++) {
  $bk = [string]$arr[$r,1]; if (-not $bk) { continue }
  $fundLines.Add((@($bk, [string]$arr[$r,2], [string]$arr[$r,3]) -join ","))
}

if ($wbOpened) { $wb.Close($false) }
if ($opened) { $xl.Quit() }

# ── 5) POST ──────────────────────────────────────────────────────────────────
if (($riskLines.Count - 1) -lt 5) { Write-Host "[upload] GUARD: 데이터 너무 적음 — 업로드 중단"; exit 3 }
$payload = @{
  risk  = ($riskLines -join "`n")
  bonds = ($bondLines -join "`n")
  carry = ($carryLines -join "`n")
  funds = ($fundLines -join "`n")
} | ConvertTo-Json -Compress
$bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
$headers = @{ "x-import-token" = $TOKEN }
try {
  $resp = Invoke-RestMethod -Method Post -Uri $ENDPOINT -Body $bytes -ContentType "application/json; charset=utf-8" -Headers $headers -TimeoutSec 300
  Write-Host ("[upload] OK — asOf " + $resp.asOf + " / risk " + $resp.risk + " / carry " + $resp.carry + " / holdings " + $resp.holdings)
} catch {
  Write-Host ("[upload] FAIL: " + $_.Exception.Message)
  exit 1
}
