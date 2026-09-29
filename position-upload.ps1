# position-upload.ps1 — 회사(K-Bond) PC용 단일 파일: 포지션 엑셀 추출 → FICC 모니터 API로 업로드.
#  파일 반출 없이 회사 PC에서 실행 (K-Bond 수집기와 같은 아웃바운드 HTTP 패턴).
#  사용: powershell -ExecutionPolicy Bypass -File position-upload.ps1 [-LatestOnly]   (또는 position-upload-now.exe 더블클릭)
#  ※ 이 파일은 UTF-8 BOM 인코딩 유지 필수 (PS 5.1 한글).
#
#  동작(2026-09-29 개정):
#   · 폴더의 App.금리차익.포지션.YYYY.MM.DD*.xlsm 을 파일명 날짜 오름차순으로 처리.
#     - 최신 파일: 매 실행 업로드(장중 갱신 반영, 종전과 동일)
#     - 과거 파일: 옆에 "<파일명>.uploaded" 마커가 없으면 1회 업로드 후 마커 생성 → 휴가 등으로 밀린 기간을
#       폴더에 넣어두기만 하면 다음 실행(스케줄/즉시 exe)에서 자동 백필. -LatestOnly 면 최신 1개만.
#     - 서버는 시트 기준일(asOf) 단위로 교체 저장, 보유종목은 마지막(최신) 파일로 남는다.
#   · SIRS 단기금리(RP·회사금리)는 고정 좌표(T29/T30)가 아니라 S열 라벨("RP금리"/"회사금리")로 행을 찾는다.
#     (2026-09-03 시트에 헤더 행이 삽입돼 한 줄씩 밀리면서 SOFR_3M/RP가 RP/회사로 올라가던 사고 재발 방지)
#   · 토큰: 스크립트 옆 position-upload.token.txt(구버전 이름) 또는 import-token.txt 를 우선 사용(스크립트 갱신 시 토큰 보존).
param([switch]$LatestOnly)
$ErrorActionPreference = "Continue"

# ── 설정 (회사 PC 환경에 맞게) ───────────────────────────────────────────────
$ENDPOINT = "https://bondmonitoring.onrender.com/api/import/positions"
$TOKEN    = "CHANGE_ME"                                  # ★ FICC 모니터 IMPORT_TOKEN (또는 import-token.txt 사용)
$FILE_DIR = "C:\kbond-collector\positions"               # App.금리차익.포지션.*.xlsm 을 넣어두는 폴더
$BOOK     = "B020105"                                    # 금리차익 북

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
# ── 동시 실행 방지 — 정기 태스크(매시)와 수동 백필이 겹치면 같은 파일을 두 번 올리고(서버 적재 ~100초) 같은 asOf 를
#    동시에 지우고 써서 서버 오류로 중단될 수 있다. 락 파일(30분 이내 생성)이 있으면 이번 실행은 건너뛴다.
$lockFile = Join-Path $scriptDir "position-upload.lock"
if (Test-Path $lockFile) {
  $age = (Get-Date) - (Get-Item $lockFile).LastWriteTime
  if ($age.TotalMinutes -lt 30) { Write-Host ("[upload] 다른 업로드 실행 중(락 " + [int]$age.TotalMinutes + "분 전) — 이번 실행 건너뜀"); exit 0 }
  Write-Host "[upload] 오래된 락(30분 초과) 무시"
}
Set-Content -Path $lockFile -Value ("pid " + $PID + " " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Encoding ASCII
# 토큰파일: 구버전 이름(position-upload.token.txt) 우선, 없으면 import-token.txt — 둘 중 하나만 있으면 됨
$tokenFile = $null
foreach ($n in @("position-upload.token.txt", "import-token.txt")) { $c = Join-Path $scriptDir $n; if (Test-Path $c) { $tokenFile = $c; break } }
if ($tokenFile) { $t = (Get-Content $tokenFile -Raw).Trim(); if ($t) { $TOKEN = $t } }
if (-not $TOKEN -or $TOKEN -eq "CHANGE_ME") { Write-Host "[upload] 토큰 없음 — $scriptDir\position-upload.token.txt 에 IMPORT_TOKEN 을 한 줄로 넣으세요"; exit 4 }

# ── 대상 파일 목록 (파일명 날짜 오름차순, 같은 날짜는 가장 늦게 저장된 버전) ─────────
$all = Get-ChildItem $FILE_DIR -File -Filter "App.금리차익.포지션.*.xlsm" -Recurse -Depth 2 -ErrorAction SilentlyContinue | ForEach-Object {
  $m = [regex]::Match($_.Name, '(\d{4})\.(\d{2})\.(\d{2})')
  if ($m.Success) { [PSCustomObject]@{ Date = ($m.Groups[1].Value + "-" + $m.Groups[2].Value + "-" + $m.Groups[3].Value); File = $_ } }
}
if (-not $all) { Write-Host "[upload] 포지션 파일 없음: $FILE_DIR"; exit 2 }
$byDate = $all | Group-Object Date | ForEach-Object { $_.Group | Sort-Object { $_.File.LastWriteTime } -Descending | Select-Object -First 1 } | Sort-Object Date
$latest = ($byDate | Select-Object -Last 1)
$targets = @()
foreach ($x in $byDate) {
  $isLatest = ($x.Date -eq $latest.Date)
  if ($LatestOnly -and -not $isLatest) { continue }
  $marker = $x.File.FullName + ".uploaded"
  if ($isLatest -or -not (Test-Path $marker)) { $targets += [PSCustomObject]@{ Date = $x.Date; File = $x.File; IsLatest = $isLatest; Marker = $marker } }
}
Write-Host ("[upload] 폴더 파일 " + @($byDate).Count + "일치, 업로드 대상 " + @($targets).Count + "개 (최신 " + $latest.Date + ")")

# ── Excel COM 1회 ────────────────────────────────────────────────────────────
$xl = $null; $opened = $false
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $opened = $true }

function Sheet($wb, $name) { foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { return $s } }; return $null }
function OaYmd($v) { if ($v -is [double] -and $v -gt 20000) { return [DateTime]::FromOADate($v).ToString("yyyy-MM-dd") }; return "" }
function Num($v) { if ($v -is [double] -or $v -is [int]) { return [string]$v }; return "" }
function AddMeta($lines, $asOf, $key, $v, $label) { if ($v -is [double] -or $v -is [int]) { $lines.Add("$asOf,$key,$v," + ($label -replace ',',' ')) } }
# S열 라벨로 행 찾기 (블록은 Range.Value2 2차원 배열, 1-based [행,열]) — 공백 무시
function FindLabelRow($blk, $nRows, $label) {
  for ($i = 1; $i -le $nRows; $i++) { $s = [string]$blk[$i,1]; if ($s -and ($s -replace '\s','') -eq $label) { return $i } }
  return 0
}

function Upload-One($f, $isLatest) {
  Write-Host ("[upload] file: " + $f.FullName)
  $wb = $null; foreach ($b in $xl.Workbooks) { if ($b.FullName -eq $f.FullName) { $wb = $b } }
  $wbOpened = $false
  if (-not $wb) { $wb = $xl.Workbooks.Open($f.FullName, 0, $true); $wbOpened = $true }
  try {
    # 1) 금리민감도
    $ms = Sheet $wb "금리민감도"
    if (-not $ms) { Write-Host "[upload] 금리민감도 시트 없음 — 스킵"; return $false }
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

    # 2) 채무증권
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

    # 3) 캐리 시나리오 요약
    $cs = Sheet $wb "캐리 시나리오"
    $arr = $cs.Range($cs.Cells(1,1), $cs.Cells(30, 30)).Value2
    $carryLines = New-Object System.Collections.Generic.List[string]
    $carryLines.Add("asOf,key,value,label")
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

    # 3.5) SIRS 단기금리 — S열 라벨("RP금리"/"회사금리")로 행 탐색, 값=T열, 전일대비=U열 (좌표 고정 금지)
    $ss = Sheet $wb "SIRS"
    if ($ss) {
      $nBlk = 40
      $blk = $ss.Range($ss.Cells(15,19), $ss.Cells(15 + $nBlk - 1, 21)).Value2   # S15:U54
      $iRp = FindLabelRow $blk $nBlk "RP금리"
      $iCo = FindLabelRow $blk $nBlk "회사금리"
      if ($iRp) { AddMeta $carryLines $asOf "rate.rp" $blk[$iRp,2] ("RP금리(SIRS S" + (14 + $iRp) + " 라벨)"); AddMeta $carryLines $asOf "rate.rp.chg" $blk[$iRp,3] "RP 전일대비(시트 U열)" }
      else { Write-Host "[upload] SIRS 'RP금리' 라벨 없음(S15:S54) — 스킵" }
      if ($iCo) { AddMeta $carryLines $asOf "rate.corp3m" $blk[$iCo,2] ("회사3M(SIRS S" + (14 + $iCo) + " 라벨)"); AddMeta $carryLines $asOf "rate.corp3m.chg" $blk[$iCo,3] "회사3M 전일대비(시트 U열)" }
      else { Write-Host "[upload] SIRS '회사금리' 라벨 없음(S15:S54) — 스킵" }
      $rpRow = if ($iRp) { 14 + $iRp } else { "-" }
      $coRow = if ($iCo) { 14 + $iCo } else { "-" }
      Write-Host ("[upload] SIRS 단기금리: RP행 " + $rpRow + " / 회사행 " + $coRow)
    } else { Write-Host "[upload] SIRS 시트 없음 — 단기금리 스킵" }

    # 4) 펀드→전략
    $fc = Sheet $wb "Query.펀드코드.공학팀"
    $nR = $fc.UsedRange.Rows.Count
    $arr = $fc.Range($fc.Cells(1,1), $fc.Cells($nR, 3)).Value2
    $fundLines = New-Object System.Collections.Generic.List[string]
    $fundLines.Add("bookCode,fundCode,strategy")
    for ($r = 2; $r -le $nR; $r++) {
      $bk = [string]$arr[$r,1]; if (-not $bk) { continue }
      $fundLines.Add((@($bk, [string]$arr[$r,2], [string]$arr[$r,3]) -join ","))
    }
  } finally {
    if ($wbOpened) { $wb.Close($false) }
  }

  # 5) POST
  if (($riskLines.Count - 1) -lt 5) { Write-Host "[upload] GUARD: 데이터 너무 적음 — 이 파일 업로드 중단"; return $false }
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
    return $true
  } catch {
    Write-Host ("[upload] FAIL: " + $_.Exception.Message)
    return $false
  }
}

$ok = 0; $fail = 0
foreach ($t in $targets) {
  Set-Content -Path $lockFile -Value ("pid " + $PID + " " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Encoding ASCII  # 락 갱신
  $tag = if ($t.IsLatest) { " (최신)" } else { " (백필)" }
  Write-Host ("`n[upload] === " + $t.Date + $tag + " ===")
  $r = Upload-One $t.File $t.IsLatest
  if ($r) { $ok++; if (-not $t.IsLatest) { Set-Content -Path $t.Marker -Value ("uploaded " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Encoding ASCII } }
  else { $fail++; if (-not $t.IsLatest) { Write-Host "[upload] 백필 파일 실패 — 이후 파일 중단(순서 보존)"; break } }
}
if ($opened) { $xl.Quit() }
Remove-Item -Path $lockFile -ErrorAction SilentlyContinue
Write-Host ("`n[upload] 완료 — 성공 " + $ok + " / 실패 " + $fail + " / 대상 " + @($targets).Count)
if ($fail -gt 0) { exit 1 }
