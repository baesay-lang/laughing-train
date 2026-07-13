# Doc2Clipboard

회사 PC 등에서 **파일 업로드가 막혀 있을 때**, 폴더 안의 문서를 텍스트로 변환해
`claude.ai` 채팅창에 **붙여넣기(Ctrl+V)** 로 넘길 수 있게 해 주는 PowerShell 스크립트입니다.

- **설치 프로그램 없음** — Windows에 기본 내장된 PowerShell + 이미 깔려 있는 MS Office만 사용합니다.
- AppLocker / WDAC 같은 실행 정책은 "미승인 exe"를 막는 것이지, PowerShell과 Office COM은 시스템 기본 도구라 걸리지 않습니다.
- 지원 형식: `docx` `doc` `pptx` `ppt` `xlsx` `xlsm` `pdf(텍스트 기반)` `txt` `md` `csv`

---

## 동작 방식

1. 폴더 경로를 입력하면 그 안의 문서를 전부 스캔합니다.
2. Word / Excel / PowerPoint를 **백그라운드**로 열어 텍스트를 뽑아 `_md_export` 하위 폴더에 `.md`로 저장합니다.
   - 엑셀: 시트별 `|` 구분 표
   - PPT: 슬라이드별 텍스트 + 표 + 발표자 노트
3. **클립보드 모드**: 각 문서를 15,000자 청크로 잘라 `[문서: 이름 | 부분 1/3]` 라벨을 붙여 클립보드에 복사합니다.
   claude.ai에 `Ctrl+V` → 스크립트 창으로 돌아와 `Enter` → 다음 청크. 이 반복이 전부입니다.

---

## 실행 방법

### 방법 A — 콘솔에 붙여넣기 (가장 확실함, 실행 정책과 무관)

`.ps1` 파일 실행이 막혀 있어도 됩니다. `Doc2Clipboard.ps1` **내용 전체를 복사**해
PowerShell 창에 그대로 붙여넣고 Enter를 누르세요. 폴더 경로를 물어봅니다.

### 방법 B — 파일로 실행

```powershell
powershell -ExecutionPolicy Bypass -File .\Doc2Clipboard.ps1
```

### 옵션

```powershell
# 폴더/청크 크기 지정
.\Doc2Clipboard.ps1 -FolderPath "C:\작업\분석대상" -ChunkSize 30000

# 변환만 하고 클립보드 모드는 건너뛰기 (.md 파일만 필요할 때)
.\Doc2Clipboard.ps1 -FolderPath "C:\작업\분석대상" -ConvertOnly

# 모든 문서를 하나로 합친 _ALL.md 도 함께 생성
.\Doc2Clipboard.ps1 -FolderPath "C:\작업\분석대상" -Combined
```

| 파라미터        | 설명                                   | 기본값  |
| --------------- | -------------------------------------- | ------- |
| `-FolderPath`   | 문서 폴더 경로 (생략 시 실행 중 물어봄) | (없음)  |
| `-ChunkSize`    | 붙여넣기 1회당 글자 수                  | 15000   |
| `-ConvertOnly`  | 변환만, 클립보드 모드 생략             | off     |
| `-Combined`     | `_ALL.md` 합본 생성                    | off     |

클립보드 모드 중 조작: `Enter` = 다음 청크 / `s` = 이 문서 건너뛰기 / `q` = 종료

---

## 한계와 보완

- **스캔본 PDF(이미지)** 는 텍스트가 없어 추출되지 않습니다.
  이런 건 `Win+Shift+S`로 화면 캡처 → claude.ai에 `Ctrl+V`가 제일 빠릅니다.
  (이미지 붙여넣기는 파일 업로드 차단과 별개로 대부분 됩니다. 도면·복잡한 표도 캡처가 더 정확합니다.)
- **텍스트 기반 PDF** 는 Word가 열어 변환합니다. Word 버전에 따라 "PDF를 편집 가능한 형식으로 변환" 안내가 뜰 수 있는데, 백그라운드에서 처리됩니다.
- 그림·다이어그램이 많은 문서는 텍스트만 넘어가니, **텍스트는 스크립트로 + 핵심 그림은 캡처로** 병행하는 게 최적입니다.
- claude.ai는 긴 붙여넣기를 자동으로 텍스트 첨부로 바꿔 주므로 `-ChunkSize`를 크게 잡아도 됩니다.

---

## ⚠️ 보안 주의

이 방식도 결국 회사 문서를 외부(Claude)로 내보내는 것입니다.
**수출통제·기밀 대상 자료(예: 10 CFR 810 해당 자료)는 절대 사용하지 마세요.**
일반 업무 문서는 사내에서 Claude 웹이 허용된 범위 안에서만 판단하여 사용하세요.
