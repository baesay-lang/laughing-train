<#
================================================================
 Setup.ps1  —  Doc2Clipboard 설치 관리자 (한 번만 실행하면 됨)
================================================================
 이 스크립트 내용을 PowerShell 창에 통째로 붙여넣고 Enter 치면:
   1) Doc2Clipboard 본체를 %USERPROFILE%\Doc2Clipboard 에 저장
   2) 바탕화면 바로가기 생성 (더블클릭 → 폴더 물어봄)
   3) 우클릭 '보내기(Send to)' 메뉴에 등록
      → 앞으로는 폴더에서 [우클릭 → 보내기 → Doc2Clipboard] 한 번이면 끝

 실행파일(.exe) 설치가 없어 AppLocker/WDAC 정책과 무관합니다.
 바로가기는 파일을 텍스트로 읽어 실행(iex)하므로, .ps1 직접 실행이
 막힌 환경에서도 '콘솔 붙여넣기'와 동일하게 동작합니다.

 제거하려면: 바탕화면/보내기 폴더의 Doc2Clipboard 바로가기와
             %USERPROFILE%\Doc2Clipboard 폴더를 삭제하세요.
================================================================
#>

$ErrorActionPreference = 'Stop'
try {

$installDir = Join-Path $env:USERPROFILE 'Doc2Clipboard'
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
$scriptPath = Join-Path $installDir 'Doc2Clipboard.ps1'

# ---- 아래 @' ... '@ 안이 Doc2Clipboard 본체입니다(그대로 저장됨) ----
$body = @'
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
# ---- 본체 끝 ----

# UTF-8(BOM 포함)로 저장 → 한글 깨짐 방지
[System.IO.File]::WriteAllText($scriptPath, $body, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "설치됨: $scriptPath" -ForegroundColor Green

$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$ws = New-Object -ComObject WScript.Shell

# 본체 파일을 텍스트로 읽어 그대로 실행. .ps1 실행정책/스크립트정책과 무관.
$readSelf = "iex ([System.IO.File]::ReadAllText('$scriptPath',[System.Text.Encoding]::UTF8))"

# 1) 바탕화면 바로가기 (더블클릭 → 폴더를 물어봄)
$desktop = [Environment]::GetFolderPath('Desktop')
$lnk = $ws.CreateShortcut((Join-Path $desktop 'Doc2Clipboard.lnk'))
$lnk.TargetPath       = $ps
$lnk.Arguments        = "-NoProfile -ExecutionPolicy Bypass -Command `"$readSelf`""
$lnk.WorkingDirectory = $installDir
$lnk.IconLocation     = "$ps,0"
$lnk.Description       = 'Doc2Clipboard - 문서를 claude.ai에 붙여넣기'
$lnk.Save()
Write-Host "바탕화면 바로가기 생성됨" -ForegroundColor Green

# 2) 우클릭 '보내기(Send to)' 메뉴 등록
#    폴더를 우클릭하면 그 폴더 경로가 마지막 인자로 넘어옴 → $args[0]
$sendto = [Environment]::GetFolderPath('SendTo')
$cmd = "`$env:D2C_FOLDER=`$args[0]; $readSelf"
$lnk2 = $ws.CreateShortcut((Join-Path $sendto 'Doc2Clipboard (→claude.ai).lnk'))
$lnk2.TargetPath       = $ps
$lnk2.Arguments        = "-NoProfile -ExecutionPolicy Bypass -Command `"$cmd`""
$lnk2.WorkingDirectory = $installDir
$lnk2.IconLocation     = "$ps,0"
$lnk2.Description       = '선택한 폴더의 문서를 변환해 claude.ai에 붙여넣기'
$lnk2.Save()
Write-Host "우클릭 '보내기' 메뉴에 등록됨" -ForegroundColor Green

Write-Host ""
Write-Host "=== 설치 완료! ===" -ForegroundColor Cyan
Write-Host "사용법 1) 문서가 든 폴더에서  [우클릭 → 보내기 → Doc2Clipboard]" -ForegroundColor Cyan
Write-Host "사용법 2) 바탕화면 [Doc2Clipboard] 아이콘 더블클릭 → 폴더 경로 입력" -ForegroundColor Cyan

}
catch {
    Write-Host ""
    Write-Host "설치 중 오류: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "COM(WScript.Shell) 생성이 정책으로 막혔을 수 있습니다. 이 메시지를 알려주세요." -ForegroundColor Yellow
}
Read-Host "창을 닫으려면 Enter를 누르세요"
