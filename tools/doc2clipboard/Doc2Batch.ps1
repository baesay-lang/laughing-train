<#
================================================================
 Doc2Batch.ps1
================================================================
 폴더 안의 문서를 claude.ai에 "붙여넣기(Ctrl+V)"로 넘기기 좋게
 변환하고, 여러 개를 한꺼번에 클립보드에 묶어서 배치(기본 20개)
 단위로 넘겨 주는 스크립트. 업로드가 막힌 PC용.

 두 가지 모드:
   [기본] 파일 배치 모드
     - Office 문서(docx/xlsx/pptx 등)는 PDF로 변환, PDF/이미지는 그대로.
     - 20개씩 클립보드에 "파일 묶음"으로 올림 → 채팅창에 Ctrl+V 한 번에 20개.
     - Claude는 PDF를 텍스트+그림(스캔본은 시각 인식)으로 읽으므로
       도표/스캔 문서도 그대로 살아납니다.

   [D2C_SINGLE=1] 낱장 이미지 모드
     - 각 페이지를 PNG로 렌더해서 "한 장씩" 클립보드(이미지)로 복사.
     - 파일 여러 개 붙여넣기가 안 되는 브라우저용 확실한 대비책.

 필요한 것 : Windows PowerShell 5.1 + MS Office (설치 프로그램 없음)
 설정(환경변수):
   D2C_FOLDER   변환할 폴더 (없으면 물어봄)
   D2C_BATCH    한 배치당 파일 수 (기본 20)   ← 파일 배치 모드
   D2C_SINGLE   값이 있으면 낱장 이미지 모드
   D2C_SCALE    PNG 렌더 배율 (기본 2.0)      ← 낱장 이미지 모드
================================================================
#>

trap {
    Write-Host ""
    Write-Host "오류가 발생했습니다: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "창을 닫으려면 Enter를 누르세요"
    break
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- 설정 ----------
$FolderPath = if ($env:D2C_FOLDER) { $env:D2C_FOLDER } else { "" }
$Batch      = if ($env:D2C_BATCH)  { [int]$env:D2C_BATCH } else { 20 }
$Single     = [bool]$env:D2C_SINGLE
$Scale      = if ($env:D2C_SCALE)  { [double]$env:D2C_SCALE } else { 2.0 }

if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    $FolderPath = Read-Host "문서가 들어있는 폴더 경로를 입력하세요"
}
$FolderPath = $FolderPath.Trim().Trim('"').Trim()
if (-not (Test-Path -LiteralPath $FolderPath)) {
    Write-Host "폴더를 찾을 수 없습니다: $FolderPath" -ForegroundColor Red
    Read-Host "창을 닫으려면 Enter를 누르세요"; return
}

$outDir = Join-Path $FolderPath "_chat_export"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Release-Com { param($o) if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

# ---------- Office → PDF ----------
function Convert-ToPdf {
    param([string]$Path, [string]$PdfOut, [string]$Ext)
    switch -Regex ($Ext) {
        '\.(docx|doc|rtf)$' {
            $app=$null;$doc=$null
            try { $app=New-Object -ComObject Word.Application } catch { return "WORD_COM_BLOCKED" }
            try {
                $app.Visible=$false; $app.DisplayAlerts=0
                $doc=$app.Documents.Open($Path,$false,$true)
                $doc.ExportAsFixedFormat($PdfOut,17)   # 17 = wdExportFormatPDF
                return "OK"
            } catch { return "FAIL: $($_.Exception.Message)" }
            finally { if($doc){try{$doc.Close($false)}catch{}}; if($app){try{$app.Quit()}catch{}}; Release-Com $doc; Release-Com $app; [GC]::Collect() }
        }
        '\.(xlsx|xlsm|xls|csv)$' {
            $app=$null;$wb=$null
            try { $app=New-Object -ComObject Excel.Application } catch { return "EXCEL_COM_BLOCKED" }
            try {
                $app.Visible=$false; $app.DisplayAlerts=$false
                $wb=$app.Workbooks.Open($Path,0,$true)
                $wb.ExportAsFixedFormat(0,$PdfOut)     # 0 = xlTypePDF
                return "OK"
            } catch { return "FAIL: $($_.Exception.Message)" }
            finally { if($wb){try{$wb.Close($false)}catch{}}; if($app){try{$app.Quit()}catch{}}; Release-Com $wb; Release-Com $app; [GC]::Collect() }
        }
        '\.(pptx|ppt)$' {
            $app=$null;$pres=$null
            try { $app=New-Object -ComObject PowerPoint.Application } catch { return "PPT_COM_BLOCKED" }
            try {
                $pres=$app.Presentations.Open($Path,$true,$false,$false)
                $pres.SaveAs($PdfOut,32)               # 32 = ppSaveAsPDF
                return "OK"
            } catch { return "FAIL: $($_.Exception.Message)" }
            finally { if($pres){try{$pres.Close()}catch{}}; if($app){try{$app.Quit()}catch{}}; Release-Com $pres; Release-Com $app; [GC]::Collect() }
        }
        default { return "SKIP" }
    }
}

# ---------- WinRT: PDF → PNG (낱장 이미지 모드용) ----------
$script:winrtReady = $false
function Init-WinRT {
    if ($script:winrtReady) { return $true }
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        $null=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
        $null=[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]
        $null=[Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
        $script:winrtReady=$true; return $true
    } catch { return $false }
}
function Await-Op($op,$t) {
    $m=[System.WindowsRuntimeSystemExtensions].GetMethods()|Where-Object{$_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'}|Select-Object -First 1
    $task=$m.MakeGenericMethod($t).Invoke($null,@($op)); $task.Wait(-1)|Out-Null; $task.Result
}
function Await-Act($act) {
    $m=[System.WindowsRuntimeSystemExtensions].GetMethods()|Where-Object{$_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'}|Select-Object -First 1
    $task=$m.Invoke($null,@($act)); $task.Wait(-1)|Out-Null
}
function Render-PdfToPng {
    param([string]$PdfPath,[string]$OutDir,[string]$BaseName)
    $pngs=@()
    $sf   = Await-Op ([Windows.Storage.StorageFile]::GetFileFromPathAsync($PdfPath)) ([Windows.Storage.StorageFile])
    $pdf  = Await-Op ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
    $folder = Await-Op ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($OutDir)) ([Windows.Storage.StorageFolder])
    for ($i=0; $i -lt $pdf.PageCount; $i++) {
        $page=$pdf.GetPage($i)
        $pname=("{0}_p{1:D3}.png" -f $BaseName,($i+1))
        $of  = Await-Op ($folder.CreateFileAsync($pname,1)) ([Windows.Storage.StorageFile])   # 1=ReplaceExisting
        $stream = Await-Op ($of.OpenAsync(1)) ([Windows.Storage.Streams.IRandomAccessStream]) # 1=ReadWrite
        $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $opt.DestinationWidth = [uint32]([math]::Round($page.Size.Width * $Scale))
        Await-Act ($page.RenderToStreamAsync($stream,$opt))
        $stream.Dispose(); $page.Dispose()
        $pngs += (Join-Path $OutDir $pname)
    }
    return $pngs
}
# PPT는 WinRT 없이 네이티브로 PNG 내보내기 가능
function Export-PptToPng {
    param([string]$Path,[string]$OutDir,[string]$BaseName)
    $app=$null;$pres=$null
    try { $app=New-Object -ComObject PowerPoint.Application } catch { return @() }
    try {
        $pres=$app.Presentations.Open($Path,$true,$false,$false)
        $sub=Join-Path $OutDir $BaseName; New-Item -ItemType Directory -Force -Path $sub|Out-Null
        $pres.SaveAs($sub,18)    # 18 = ppSaveAsPNG (슬라이드별 PNG 폴더)
        return (Get-ChildItem -LiteralPath $sub -Filter *.PNG -Recurse | Sort-Object Name | ForEach-Object { $_.FullName })
    } catch { return @() }
    finally { if($pres){try{$pres.Close()}catch{}}; if($app){try{$app.Quit()}catch{}}; Release-Com $pres; Release-Com $app; [GC]::Collect() }
}

# ---------- 대상 스캔 ----------
$docExt = '\.(docx|doc|rtf|xlsx|xlsm|xls|csv|pptx|ppt)$'
$imgExt = '\.(png|jpg|jpeg|gif|bmp)$'
$files = Get-ChildItem -LiteralPath $FolderPath -File | Where-Object {
    ($_.Extension -match $docExt -or $_.Extension -match $imgExt -or $_.Extension -match '\.pdf$') -and $_.Name -notlike '~$*'
} | Sort-Object Name
if ($files.Count -eq 0) { Write-Host "변환할 문서가 없습니다." -ForegroundColor Yellow; Read-Host "Enter"; return }

Write-Host ""
Write-Host ("=== 대상 {0}개 / 모드: {1} ===" -f $files.Count, $(if($Single){"낱장 이미지"}else{"파일 배치(PDF)"})) -ForegroundColor Cyan

$blocked = $false
$outFiles = @()   # 배치로 넘길 최종 파일 목록

foreach ($f in $files) {
    Write-Host ("처리 중: {0} ..." -f $f.Name) -NoNewline
    $base = ($f.BaseName -replace '[^\w\.\- ]','_')

    if (-not $Single) {
        # ---- 파일 배치 모드: PDF/이미지 파일로 준비 ----
        if ($f.Extension -match $imgExt -or $f.Extension -match '\.pdf$') {
            $outFiles += $f.FullName; Write-Host " OK" -ForegroundColor Green; continue
        }
        $pdf = Join-Path $outDir ($base + ".pdf")
        $r = Convert-ToPdf -Path $f.FullName -PdfOut $pdf -Ext $f.Extension
        if ($r -eq "OK") { $outFiles += $pdf; Write-Host " PDF" -ForegroundColor Green }
        elseif ($r -like "*_COM_BLOCKED") { $blocked=$true; Write-Host " 차단됨($r)" -ForegroundColor Red }
        else { Write-Host " 실패($r)" -ForegroundColor Yellow }
    }
    else {
        # ---- 낱장 이미지 모드: PNG로 렌더 ----
        if ($f.Extension -match $imgExt) { $outFiles += $f.FullName; Write-Host " OK" -ForegroundColor Green; continue }
        if ($f.Extension -match '\.(pptx|ppt)$') {
            $pngs = Export-PptToPng -Path $f.FullName -OutDir $outDir -BaseName $base
            if ($pngs.Count) { $outFiles += $pngs; Write-Host (" PNG x{0}" -f $pngs.Count) -ForegroundColor Green }
            else { $blocked=$true; Write-Host " 실패(PPT COM)" -ForegroundColor Red }
            continue
        }
        # docx/xlsx/pdf → (PDF) → WinRT PNG
        if (-not (Init-WinRT)) { Write-Host " 실패(WinRT PDF 렌더 불가)" -ForegroundColor Red; $blocked=$true; continue }
        $srcPdf = $f.FullName
        if ($f.Extension -notmatch '\.pdf$') {
            $srcPdf = Join-Path $outDir ($base + ".pdf")
            $r = Convert-ToPdf -Path $f.FullName -PdfOut $srcPdf -Ext $f.Extension
            if ($r -ne "OK") { if($r -like "*_COM_BLOCKED"){$blocked=$true}; Write-Host " 실패($r)" -ForegroundColor Yellow; continue }
        }
        try {
            $pngs = Render-PdfToPng -PdfPath $srcPdf -OutDir $outDir -BaseName $base
            $outFiles += $pngs; Write-Host (" PNG x{0}" -f $pngs.Count) -ForegroundColor Green
        } catch { Write-Host " 실패(렌더 $($_.Exception.Message))" -ForegroundColor Yellow }
    }
}

if ($outFiles.Count -eq 0) {
    Write-Host ""
    if ($blocked) {
        Write-Host "Office/렌더 기능이 정책으로 막혀 변환이 안 됩니다." -ForegroundColor Red
        Write-Host "→ 이 경우 자동 변환은 불가합니다. 화면 캡처(Win+Shift+S) 방식으로 넘겨야 합니다." -ForegroundColor Yellow
    } else {
        Write-Host "변환된 결과가 없습니다." -ForegroundColor Yellow
    }
    Read-Host "창을 닫으려면 Enter를 누르세요"; return
}

Write-Host ""
Write-Host ("준비 완료: {0}개 항목  (저장: {1})" -f $outFiles.Count, $outDir) -ForegroundColor Cyan
if ($blocked) { Write-Host "일부 문서는 정책 차단으로 건너뛰었습니다." -ForegroundColor Yellow }

# ---------- 전송 ----------
if (-not $Single) {
    # 파일 배치: 20개씩 클립보드에 파일 묶음으로 올림
    Write-Host ""
    Write-Host "=== 파일 배치 전송 (한 배치 $Batch개) ===" -ForegroundColor Cyan
    Write-Host "각 배치가 클립보드에 '파일 묶음'으로 올라갑니다. 채팅창에 Ctrl+V 한 번 → 배치 전체 첨부." -ForegroundColor Gray
    Write-Host "  Enter = 다음 배치 / q = 종료" -ForegroundColor DarkGray
    Write-Host ""
    $total = [math]::Ceiling($outFiles.Count / $Batch)
    for ($b=0; $b -lt $total; $b++) {
        $slice = $outFiles[($b*$Batch)..([math]::Min(($b+1)*$Batch-1, $outFiles.Count-1))]
        $col = New-Object System.Collections.Specialized.StringCollection
        $col.AddRange([string[]]$slice)
        [System.Windows.Forms.Clipboard]::SetFileDropList($col)
        Write-Host ("-> 배치 {0}/{1}: {2}개 파일 클립보드에 올림" -f ($b+1),$total,$slice.Count) -ForegroundColor Yellow
        $ans = Read-Host "   Ctrl+V로 붙여넣은 뒤 Enter (q=종료)"
        if ($ans -eq 'q') { break }
    }
    Write-Host ""
    Write-Host "만약 Ctrl+V로 파일이 안 붙었다면: 이 브라우저가 '파일 붙여넣기'를 막는 것이라" -ForegroundColor Yellow
    Write-Host "낱장 이미지 모드가 필요합니다.  실행 전에  `$env:D2C_SINGLE=1  을 설정하고 다시 돌리세요." -ForegroundColor Yellow
}
else {
    # 낱장 이미지: 한 장씩 클립보드(이미지)로
    Write-Host ""
    Write-Host "=== 낱장 이미지 전송 ($($outFiles.Count)장) ===" -ForegroundColor Cyan
    Write-Host "이미지가 한 장씩 클립보드에 복사됩니다. 채팅창에 Ctrl+V → Enter → 다음 장." -ForegroundColor Gray
    Write-Host "  Enter = 다음 장 / q = 종료" -ForegroundColor DarkGray
    Write-Host ""
    for ($k=0; $k -lt $outFiles.Count; $k++) {
        try {
            $img=[System.Drawing.Image]::FromFile($outFiles[$k])
            [System.Windows.Forms.Clipboard]::SetImage($img)
            $img.Dispose()
        } catch { Write-Host ("   (건너뜀: {0})" -f $outFiles[$k]) -ForegroundColor DarkYellow; continue }
        Write-Host ("-> {0}/{1} 클립보드에 이미지 복사됨" -f ($k+1),$outFiles.Count) -ForegroundColor Yellow
        $ans = Read-Host "   Ctrl+V로 붙여넣은 뒤 Enter (q=종료)"
        if ($ans -eq 'q') { break }
    }
}

Write-Host ""
Write-Host "완료!" -ForegroundColor Green
Read-Host "창을 닫으려면 Enter를 누르세요"
