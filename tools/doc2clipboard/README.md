# Doc2Clipboard

파일 업로드가 막힌 PC에서, 폴더 안의 문서를 텍스트로 변환해 `claude.ai` 채팅창에
**붙여넣기(Ctrl+V)** 로 넘길 수 있게 해 주는 도구입니다.

- **설치 프로그램(.exe) 없음** — Windows 기본 PowerShell + 이미 깔린 MS Office만 사용.
  미승인 exe를 막는 AppLocker / WDAC 정책과 무관하게 동작합니다.
- 지원 형식: `docx` `doc` `pptx` `ppt` `xlsx` `xlsm` `pdf(텍스트 기반)` `txt` `md` `csv`

---

## ⭐ 설치 (딱 한 번) → 이후 우클릭 한 번

매번 코드를 붙여넣지 않으려면 **`Setup.ps1`** 을 한 번만 실행하세요.

1. GitHub에서 [`Setup.ps1`](./Setup.ps1) 을 열고 **`Raw`** → `Ctrl+A` → `Ctrl+C`
2. `Win` → `powershell` → **Windows PowerShell** 실행
3. `Ctrl+V` 로 붙여넣고 `Enter`

설치되는 것:

| 항목 | 위치 | 사용법 |
| --- | --- | --- |
| 본체 스크립트 | `%USERPROFILE%\Doc2Clipboard\Doc2Clipboard.ps1` | (자동) |
| **우클릭 메뉴** | 보내기(Send to) | 폴더 **우클릭 → 보내기 → Doc2Clipboard** |
| 바탕화면 아이콘 | Desktop | 더블클릭 → 폴더 경로 입력 |

설치 후에는 **분석할 문서가 든 폴더에서 우클릭 → 보내기 → Doc2Clipboard** 만 하면 됩니다.
바로가기는 본체를 텍스트로 읽어 실행(`iex`)하므로, `.ps1` 직접 실행이 정책으로 막힌
환경에서도 "콘솔 붙여넣기"와 동일하게 동작합니다.

> 제거: 바탕화면·보내기 폴더의 `Doc2Clipboard` 바로가기와 `%USERPROFILE%\Doc2Clipboard` 폴더 삭제.

---

## 실행하면 벌어지는 일

1. (우클릭으로 넘긴) 폴더 안 문서를 전부 스캔 → Word/Excel/PowerPoint를 **백그라운드**로 열어
   텍스트를 뽑아 `_md_export` 폴더에 `.md`로 저장합니다.
   - 엑셀: 시트별 `|` 구분 표 / PPT: 슬라이드별 텍스트 + 표 + 발표자 노트
2. **클립보드 모드**: 각 문서를 청크로 잘라 `[문서: 이름 | 부분 1/3]` 라벨을 붙여 클립보드에 복사.
   claude.ai에 `Ctrl+V` → 스크립트 창으로 돌아와 `Enter` → 다음 청크. 이 반복이 전부입니다.

조작키: `Enter` = 다음 청크 / `s` = 이 문서 건너뛰기 / `q` = 종료

---

## 설치 없이 한 번만 쓰기 (대안)

바로 한 번만 돌려보고 싶으면, [`Doc2Clipboard.ps1`](./Doc2Clipboard.ps1) 내용 전체를
PowerShell 창에 붙여넣고 `Enter` → 폴더 경로를 물어봅니다. (콘솔 붙여넣기는 실행 정책과 무관)

### 설정 바꾸기 (환경변수)

우클릭/바로가기 실행에도 적용됩니다. 붙여넣기 전에 먼저 설정하면 됩니다.

```powershell
$env:D2C_CHUNKSIZE = 30000   # 붙여넣기 1회당 글자 수 (기본 15000)
$env:D2C_CONVERTONLY = 1     # 변환만, 클립보드 모드 생략 (.md 파일만 필요할 때)
$env:D2C_COMBINED = 1        # 모든 문서를 합친 _ALL.md 도 생성
```

| 환경변수 | 설명 | 기본값 |
| --- | --- | --- |
| `D2C_FOLDER` | 문서 폴더 경로 (없으면 실행 중 물어봄) | (없음) |
| `D2C_CHUNKSIZE` | 붙여넣기 1회당 글자 수 | 15000 |
| `D2C_CONVERTONLY` | 값 있으면 변환만 | off |
| `D2C_COMBINED` | 값 있으면 `_ALL.md` 합본 생성 | off |

---

## 알아두면 좋은 점

- **스캔본 PDF(이미지)·도면·복잡한 표**는 텍스트가 안 뽑힙니다. 이런 건 `Win+Shift+S`로 캡처 →
  claude.ai에 `Ctrl+V`가 제일 빠르고 정확합니다. (이미지 붙여넣기는 대부분 됩니다.)
- 첫 실행 때 Word/Excel/PowerPoint가 잠깐 백그라운드로 떴다 사라집니다 — 정상입니다.
- claude.ai는 긴 붙여넣기를 자동으로 텍스트 첨부로 바꿔 주므로 `D2C_CHUNKSIZE`를 크게 잡아도 됩니다.
