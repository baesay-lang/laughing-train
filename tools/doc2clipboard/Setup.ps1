<#
================================================================
 Setup.ps1  —  문서→claude.ai 붙여넣기 도구 설치 (한 번만 실행)
================================================================
 이 내용을 PowerShell 창에 통째로 붙여넣고 Enter 치면:
   1) 도구 2개를 %USERPROFILE%\Doc2Clipboard 에 저장
        - Doc2Batch.ps1     : 문서를 PDF/이미지로 변환해 배치로 붙여넣기 (추천)
        - Doc2Clipboard.ps1 : 문서를 텍스트(.md)로 뽑아 붙여넣기
   2) 우클릭 '보내기(Send to)' 메뉴에 등록
        - Doc2Batch (추천)        : 폴더 문서를 PDF로 변환, 20개씩 파일 묶음 붙여넣기
        - Doc2Batch 낱장이미지     : 각 페이지를 PNG로, 한 장씩 붙여넣기
        - Doc2Clipboard 텍스트     : 텍스트만 뽑아 붙여넣기
   3) 바탕화면 아이콘(Doc2Batch)

 실행파일(.exe) 설치가 없어 AppLocker/WDAC 정책과 무관합니다.
 바로가기는 스크립트를 텍스트로 읽어 실행(iex)하므로 .ps1 실행 제약과도 무관합니다.

 제거: 바탕화면·보내기 폴더의 해당 바로가기와 %USERPROFILE%\Doc2Clipboard 폴더 삭제.
================================================================
#>

$ErrorActionPreference = 'Stop'
try {

$installDir = Join-Path $env:USERPROFILE 'Doc2Clipboard'
New-Item -ItemType Directory -Force -Path $installDir | Out-Null

$pathBatch = Join-Path $installDir 'Doc2Batch.ps1'
$pathText  = Join-Path $installDir 'Doc2Clipboard.ps1'
$enc = New-Object System.Text.UTF8Encoding($true)

# ==== Doc2Batch.ps1 본체 ====
$bodyBatch = @'
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
'@
[System.IO.File]::WriteAllText($pathBatch, $bodyBatch, $enc)
Write-Host "설치됨: $pathBatch" -ForegroundColor Green

# ==== Doc2Clipboard.ps1 (텍스트) 본체 ====
$bodyText = @'
<#
================================================================
 Doc2Clipboard.ps1
================================================================
 폴더 안의 문서(docx/doc/pptx/ppt/xlsx/xlsm/pdf/txt/md/csv)를
 텍스트(.md)로 변환하고, claude.ai 채팅창에 붙여넣기 좋게
 클립보드로 청크 단위 복사해 주는 스크립트.

 필요한 것 : Windows PowerShell 5.1 이상 + MS Office
             (별도 설치 프로그램 없음 → AppLocker/WDAC 정책과 무관)

 설정은 환경변수로 받습니다(우클릭 메뉴/바로가기가 자동으로 넣어줌):
   D2C_FOLDER      변환할 폴더 경로 (없으면 실행 중 물어봄)
   D2C_CHUNKSIZE   붙여넣기 1회당 글자 수 (기본 15000)
   D2C_CONVERTONLY 값이 있으면 변환만 하고 클립보드 모드 생략
   D2C_COMBINED    값이 있으면 _ALL.md 합본도 생성

 실행 방법은 같은 폴더의 README.md 및 Setup.ps1 참고.
================================================================
#>

# 창이 바로 닫히지 않도록: 예기치 못한 오류도 잡아서 멈춤
trap {
    Write-Host ""
    Write-Host "오류가 발생했습니다: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "창을 닫으려면 Enter를 누르세요"
    break
}

# ---------- 설정(환경변수 → 기본값) ----------

$FolderPath  = if ($env:D2C_FOLDER)    { $env:D2C_FOLDER }         else { "" }
$ChunkSize   = if ($env:D2C_CHUNKSIZE) { [int]$env:D2C_CHUNKSIZE } else { 15000 }
$ConvertOnly = [bool]$env:D2C_CONVERTONLY
$Combined    = [bool]$env:D2C_COMBINED

if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    $FolderPath = Read-Host "문서가 들어있는 폴더 경로를 입력하세요 (예: C:\Users\me\Documents\분석대상)"
}
$FolderPath = $FolderPath.Trim().Trim('"').Trim()

if (-not (Test-Path -LiteralPath $FolderPath)) {
    Write-Host "폴더를 찾을 수 없습니다: $FolderPath" -ForegroundColor Red
    Read-Host "창을 닫으려면 Enter를 누르세요"
    return
}

$outDir = Join-Path $FolderPath "_md_export"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# ---------- 변환 함수 ----------

function Release-Com {
    param($obj)
    if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

function Convert-WordToText {
    param([string]$Path)
    $word = $null; $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        return "[변환 실패: Word를 실행할 수 없습니다. MS Word가 설치되어 있는지 확인하세요.]"
    }
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        # Open(FileName, ConfirmConversions=$false, ReadOnly=$true)
        $doc  = $word.Documents.Open($Path, $false, $true)
        $text = $doc.Content.Text
        return $text
    }
    catch {
        return "[변환 실패: $($_.Exception.Message)]"
    }
    finally {
        if ($doc)  { try { $doc.Close($false) } catch {} }
        if ($word) { try { $word.Quit() }      catch {} }
        Release-Com $doc; Release-Com $word
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Convert-ExcelToText {
    param([string]$Path)
    $excel = $null; $wb = $null
    $sb = New-Object System.Text.StringBuilder
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        return "[변환 실패: Excel을 실행할 수 없습니다. MS Excel이 설치되어 있는지 확인하세요.]"
    }
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        # Open(FileName, UpdateLinks=0, ReadOnly=$true)
        $wb = $excel.Workbooks.Open($Path, 0, $true)
        foreach ($ws in $wb.Worksheets) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("## [시트] $($ws.Name)")
            $used = $ws.UsedRange
            $vals = $used.Value2
            if ($null -eq $vals) { continue }
            # 단일 셀이면 2차원 배열이 아님
            if ($vals -isnot [object[,]]) {
                [void]$sb.AppendLine("$vals")
                continue
            }
            $rows = $vals.GetLength(0)
            $cols = $vals.GetLength(1)
            for ($r = 1; $r -le $rows; $r++) {
                $line = @()
                for ($c = 1; $c -le $cols; $c++) {
                    $v = $vals[$r, $c]
                    $line += if ($null -eq $v) { "" } else { "$v" }
                }
                if (($line -join "").Trim() -ne "") {
                    [void]$sb.AppendLine(($line -join " | "))
                }
            }
        }
        return $sb.ToString()
    }
    catch {
        return "[변환 실패: $($_.Exception.Message)]"
    }
    finally {
        if ($wb)    { try { $wb.Close($false) } catch {} }
        if ($excel) { try { $excel.Quit() }    catch {} }
        Release-Com $wb; Release-Com $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Convert-PptToText {
    param([string]$Path)
    $ppt = $null; $pres = $null
    $sb = New-Object System.Text.StringBuilder
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
    } catch {
        return "[변환 실패: PowerPoint를 실행할 수 없습니다. MS PowerPoint가 설치되어 있는지 확인하세요.]"
    }
    try {
        # Open(FileName, ReadOnly=$true, Untitled=$false, WithWindow=$false)
        $pres = $ppt.Presentations.Open($Path, $true, $false, $false)
        $i = 1
        foreach ($slide in $pres.Slides) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("## [슬라이드 $i]")
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    [void]$sb.AppendLine($shape.TextFrame.TextRange.Text)
                }
                if ($shape.HasTable) {
                    foreach ($row in $shape.Table.Rows) {
                        $cells = @()
                        foreach ($cell in $row.Cells) {
                            $cells += $cell.Shape.TextFrame.TextRange.Text
                        }
                        [void]$sb.AppendLine(($cells -join " | "))
                    }
                }
            }
            # 발표자 노트
            if ($slide.NotesPage.Shapes.Count -ge 2) {
                $note = $slide.NotesPage.Shapes |
                        Where-Object { $_.HasTextFrame -and $_.TextFrame.HasText } |
                        Select-Object -Last 1
                if ($note) {
                    $noteText = $note.TextFrame.TextRange.Text.Trim()
                    if ($noteText -ne "") { [void]$sb.AppendLine("[노트] $noteText") }
                }
            }
            $i++
        }
        return $sb.ToString()
    }
    catch {
        return "[변환 실패: $($_.Exception.Message)]"
    }
    finally {
        if ($pres) { try { $pres.Close() } catch {} }
        if ($ppt)  { try { $ppt.Quit() }  catch {} }
        Release-Com $pres; Release-Com $ppt
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---------- 폴더 스캔 및 변환 ----------

$targets = Get-ChildItem -LiteralPath $FolderPath -File | Where-Object {
    $_.Extension -match '\.(docx|doc|pdf|xlsx|xlsm|pptx|ppt|txt|md|csv)$' -and $_.Name -notlike '~$*'
}

if ($targets.Count -eq 0) {
    Write-Host "변환할 문서가 없습니다. ($FolderPath)" -ForegroundColor Yellow
    Read-Host "창을 닫으려면 Enter를 누르세요"
    return
}

Write-Host ""
Write-Host "=== 변환 대상: $($targets.Count)개 파일 ===" -ForegroundColor Cyan
$mdFiles = @()

foreach ($f in $targets) {
    Write-Host ("변환 중: {0} ..." -f $f.Name) -NoNewline
    $text = switch -Regex ($f.Extension) {
        '\.(docx|doc|pdf)$' { Convert-WordToText  -Path $f.FullName }
        '\.(xlsx|xlsm)$'    { Convert-ExcelToText -Path $f.FullName }
        '\.(pptx|ppt)$'     { Convert-PptToText   -Path $f.FullName }
        '\.(txt|md|csv)$'   { Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 }
        default             { "" }
    }

    $header = "# 문서: $($f.Name)`n"
    $header += "(원본 형식: $($f.Extension), 크기: $([math]::Round($f.Length/1KB,1)) KB, 수정일: $($f.LastWriteTime.ToString('yyyy-MM-dd')))`n`n"

    $mdPath = Join-Path $outDir ($f.BaseName + ".md")
    ($header + $text) | Out-File -FilePath $mdPath -Encoding UTF8
    $mdFiles += $mdPath
    Write-Host " 완료" -ForegroundColor Green
}

# 합본 옵션
if ($Combined) {
    $allPath = Join-Path $outDir "_ALL.md"
    $sbAll = New-Object System.Text.StringBuilder
    foreach ($mdPath in $mdFiles) {
        [void]$sbAll.AppendLine((Get-Content -LiteralPath $mdPath -Raw -Encoding UTF8))
        [void]$sbAll.AppendLine("`n---`n")
    }
    $sbAll.ToString() | Out-File -FilePath $allPath -Encoding UTF8
    Write-Host "합본 생성: $allPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "MD 파일 저장 위치: $outDir" -ForegroundColor Cyan
try { Invoke-Item -LiteralPath $outDir } catch {}

if ($ConvertOnly) {
    Write-Host "변환만 수행하고 종료합니다 (D2C_CONVERTONLY)." -ForegroundColor Cyan
    Read-Host "창을 닫으려면 Enter를 누르세요"
    return
}

# ---------- 클립보드 청크 복사 모드 ----------

Write-Host ""
Write-Host "=== 클립보드 복사 모드 ===" -ForegroundColor Cyan
Write-Host "각 청크가 클립보드에 복사됩니다. claude.ai 채팅창에 Ctrl+V로 붙여넣은 뒤"
Write-Host "이 창으로 돌아와 Enter를 누르면 다음 청크가 복사됩니다."
Write-Host "  Enter = 다음 청크 / s = 이 문서 건너뛰기 / q = 종료" -ForegroundColor DarkGray
Write-Host ""

:docLoop foreach ($mdPath in $mdFiles) {
    $name    = [System.IO.Path]::GetFileNameWithoutExtension($mdPath)
    $content = Get-Content -LiteralPath $mdPath -Raw -Encoding UTF8
    if ([string]::IsNullOrEmpty($content)) { continue }

    # 청크 분할
    $chunks = @()
    for ($i = 0; $i -lt $content.Length; $i += $ChunkSize) {
        $len = [Math]::Min($ChunkSize, $content.Length - $i)
        $chunks += $content.Substring($i, $len)
    }

    $total = $chunks.Count
    for ($j = 0; $j -lt $total; $j++) {
        $label   = "[문서: $name | 부분 $($j+1)/$total]"
        $payload = "$label`n`n$($chunks[$j])"
        if ($j -eq $total - 1) { $payload += "`n`n[이 문서의 마지막 부분입니다]" }

        Set-Clipboard -Value $payload
        Write-Host ("-> 클립보드 복사됨: {0} ({1}/{2})" -f $name, ($j+1), $total) -ForegroundColor Yellow

        $ans = Read-Host "   붙여넣은 뒤 Enter (s=이 문서 건너뛰기, q=종료)"
        if ($ans -eq 's') { break }           # 안쪽 for 종료 -> 다음 문서
        if ($ans -eq 'q') { break docLoop }   # 전체 종료
    }
}

Write-Host ""
Write-Host "모든 문서 전송 완료!" -ForegroundColor Green
Read-Host "창을 닫으려면 Enter를 누르세요"
'@
[System.IO.File]::WriteAllText($pathText, $bodyText, $enc)
Write-Host "설치됨: $pathText" -ForegroundColor Green

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$ws = New-Object -ComObject WScript.Shell

function New-Lnk {
    param([string]$LnkPath,[string]$ScriptPath,[string]$Prefix,[string]$Desc)
    $readSelf = "iex ([System.IO.File]::ReadAllText('$ScriptPath',[System.Text.Encoding]::UTF8))"
    $cmd = "$Prefix$readSelf"
    $l = $ws.CreateShortcut($LnkPath)
    $l.TargetPath       = $ps
    $l.Arguments        = "-NoProfile -ExecutionPolicy Bypass -Command `"$cmd`""
    $l.WorkingDirectory = $installDir
    $l.IconLocation     = "$ps,0"
    $l.Description       = $Desc
    $l.Save()
}

$sendto  = [Environment]::GetFolderPath('SendTo')
$desktop = [Environment]::GetFolderPath('Desktop')
$folderArg = '$env:D2C_FOLDER=$args[0]; '

# 우클릭 '보내기' 메뉴 3종
New-Lnk (Join-Path $sendto '1. Doc2Batch (추천).lnk')      $pathBatch  $folderArg                          '폴더 문서를 PDF로 변환, 20개씩 묶어 붙여넣기'
New-Lnk (Join-Path $sendto '2. Doc2Batch 낱장이미지.lnk')  $pathBatch  ($folderArg + '$env:D2C_SINGLE=1; ') '각 페이지를 PNG로 렌더해 한 장씩 붙여넣기'
New-Lnk (Join-Path $sendto '3. Doc2Clipboard 텍스트.lnk')  $pathText   $folderArg                          '문서를 텍스트(.md)로 뽑아 붙여넣기'
Write-Host "우클릭 '보내기' 메뉴 3종 등록됨" -ForegroundColor Green

# 바탕화면 아이콘(폴더는 실행 중 물어봄)
New-Lnk (Join-Path $desktop 'Doc2Batch.lnk') $pathBatch '' '문서를 claude.ai에 붙여넣기 (PDF 배치)'
Write-Host "바탕화면 아이콘 생성됨" -ForegroundColor Green

Write-Host ""
Write-Host "=== 설치 완료! ===" -ForegroundColor Cyan
Write-Host "쓰는 법: 문서 폴더에서 [우클릭 → 보내기 → 1. Doc2Batch (추천)]" -ForegroundColor Cyan
Write-Host "         → PDF로 변환되고 20개씩 클립보드에 올라갑니다. 채팅에 Ctrl+V → Enter → 다음 배치." -ForegroundColor Cyan
Write-Host "파일이 한 번에 안 붙으면 대신 [2. Doc2Batch 낱장이미지] 를 쓰세요 (한 장씩 확실히)." -ForegroundColor Cyan

}
catch {
    Write-Host ""
    Write-Host "설치 중 오류: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "COM(WScript.Shell) 생성이 막혔을 수 있습니다. 이 메시지를 알려주세요." -ForegroundColor Yellow
}
Read-Host "창을 닫으려면 Enter를 누르세요"
