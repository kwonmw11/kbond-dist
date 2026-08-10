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

# ── 1) 금리민감도 — ★ 헤더명 기반 매핑 (컬럼 삽입/삭제에 견고. 2026-07-22 "전략구분" 삭제 사고 대응) ──
$ms = Sheet $wb "금리민감도"
$asOf = OaYmd $ms.Cells.Item(1,1).Value2
$nR = $ms.UsedRange.Rows.Count
$nC = [Math]::Min($ms.UsedRange.Columns.Count, 60)
$arr = $ms.Range($ms.Cells(1,1), $ms.Cells($nR, $nC)).Value2
$col = @{}
for ($c = 1; $c -le $nC; $c++) { $h = ([string]$arr[1,$c]).Trim(); if ($h -and -not $col.ContainsKey($h)) { $col[$h] = $c } }
function C($name) { if ($col.ContainsKey($name)) { return $col[$name] } else { return 0 } }
function V($arr, $r, $c) { if ($c -ge 1) { return $arr[$r,$c] } else { return $null } }
$required = @("북코드","펀드코드","종목코드","종목명","민감도구분","종류","자산구분","YTM","수량","PV01","1D","3M","3Y","10Y","30Y")
$missing = $required | Where-Object { -not $col.ContainsKey($_) }
if ($missing) { Write-Host ("[upload] 금리민감도 필수 헤더 누락: " + ($missing -join ",") + " — 중단"); exit 3 }
$tenorNames = @("1D","3M","6M","9M","1Y","1Y6M","2Y","2Y6M","3Y","4Y","5Y","7Y","10Y","12Y","15Y","20Y","30Y")
$riskLines = New-Object System.Collections.Generic.List[string]
$riskLines.Add("asOf,sht,bookCode,fundCode,symbol,name,riskFactor,kind,assetClass,strategy,ytm,cpn,modDur,price,maturity,bondClass,rating,quantity,bookValue,carry,pv01,t1d,t3m,t6m,t9m,t1y,t18m,t2y,t30m,t3y,t4y,t5y,t7y,t10y,t12y,t15y,t20y,t30y")
for ($r = 2; $r -le $nR; $r++) {
  if ([string](V $arr $r (C "북코드")) -ne $BOOK) { continue }
  $name = ([string](V $arr $r (C "종목명"))) -replace '[",]', ' '
  $vals = @($asOf, [string]$arr[$r,1], [string](V $arr $r (C "북코드")), [string](V $arr $r (C "펀드코드")),
    [string](V $arr $r (C "종목코드")), $name,
    [string](V $arr $r (C "민감도구분")), [string](V $arr $r (C "종류")), [string](V $arr $r (C "자산구분")),
    [string](V $arr $r (C "전략구분")),
    (Num (V $arr $r (C "YTM"))), (Num (V $arr $r (C "CPN"))), (Num (V $arr $r (C "수정듀레이션"))),
    (Num (V $arr $r (C "평가단가"))), (OaYmd (V $arr $r (C "자산만기"))),
    ([string](V $arr $r (C "채무증권분류")) -replace ',',' '), ([string](V $arr $r (C "신용등급")) -replace ',',' '),
    (Num (V $arr $r (C "수량"))), (Num (V $arr $r (C "장부금액"))), (Num (V $arr $r (C "Carry"))), (Num (V $arr $r (C "PV01"))))
  foreach ($t in $tenorNames) { $vals += (Num (V $arr $r (C $t))) }
  $riskLines.Add(($vals -join ","))
}
Write-Host ("[upload] 금리민감도 " + ($riskLines.Count - 1) + "행 (asOf " + $asOf + ")")

# ── 2) 채무증권 (헤더명 기반) ────────────────────────────────────────────────
$bs = Sheet $wb "채무증권"
$nR = $bs.UsedRange.Rows.Count
$nC2 = [Math]::Min($bs.UsedRange.Columns.Count, 60)
$arr = $bs.Range($bs.Cells(1,1), $bs.Cells($nR, $nC2)).Value2
$col = @{}
for ($c = 1; $c -le $nC2; $c++) { $h = ([string]$arr[1,$c]).Trim(); if ($h -and -not $col.ContainsKey($h)) { $col[$h] = $c } }
$required2 = @("북코드","펀드코드","종목코드","종목명","수량","평가단가","YTM")
$missing2 = $required2 | Where-Object { -not $col.ContainsKey($_) }
if ($missing2) { Write-Host ("[upload] 채무증권 필수 헤더 누락: " + ($missing2 -join ",") + " — 중단"); exit 3 }
$bondLines = New-Object System.Collections.Generic.List[string]
$bondLines.Add("asOf,sht,fundCode,symbol,name,kind,riskFactor,rating,rateType,buyDate,maturity,quantity,evalPrice,modDur,ytm,carry,buyYield,cpn,cpnFreq,buyPrice,bookValue,evalValue")
for ($r = 2; $r -le $nR; $r++) {
  if ([string](V $arr $r (C "북코드")) -ne $BOOK) { continue }
  $name = ([string](V $arr $r (C "종목명"))) -replace '[",]', ' '
  $bondLines.Add((@($asOf, [string]$arr[$r,1], [string](V $arr $r (C "펀드코드")), [string](V $arr $r (C "종목코드")), $name,
    [string](V $arr $r (C "종류")), [string](V $arr $r (C "민감도구분")),
    ([string](V $arr $r (C "신용등급")) -replace ',',' '), [string](V $arr $r (C "이자유형")),
    (OaYmd (V $arr $r (C "매수일자"))), (OaYmd (V $arr $r (C "만기일자"))),
    (Num (V $arr $r (C "수량"))), (Num (V $arr $r (C "평가단가"))), (Num (V $arr $r (C "수정듀레이션"))),
    (Num (V $arr $r (C "YTM"))), (Num (V $arr $r (C "Carry"))), (Num (V $arr $r (C "매수수익률"))),
    (Num (V $arr $r (C "CPN"))), (Num (V $arr $r (C "CPN 지급주기"))), (Num (V $arr $r (C "매수단가"))),
    (Num (V $arr $r (C "장부금액"))), (Num (V $arr $r (C "평가금액")))) -join ","))
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

# ── 3.5) SIRS 시트 단기금리 — RP(T29/U29)·회사3M(T30/U30). 값=금리, 옆칸=전일대비 ──
$ss = Sheet $wb "SIRS"
if ($ss) {
  AddMeta $carryLines $asOf "rate.rp"         $ss.Cells.Item(29,20).Value2 "RP금리(SIRS T29)"
  AddMeta $carryLines $asOf "rate.rp.chg"     $ss.Cells.Item(29,21).Value2 "RP 전일대비(U29)"
  AddMeta $carryLines $asOf "rate.corp3m"     $ss.Cells.Item(30,20).Value2 "회사3M(SIRS T30)"
  AddMeta $carryLines $asOf "rate.corp3m.chg" $ss.Cells.Item(30,21).Value2 "회사3M 전일대비(U30)"
  Write-Host "[upload] SIRS 단기금리(RP/회사3M) 추출"
} else { Write-Host "[upload] SIRS 시트 없음 — 단기금리 스킵" }

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
