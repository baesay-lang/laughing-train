# WorkReport AutoFill 런처 사용법

`WorkReport_AutoFill_실행.bat`은 "Python was not found" 오류를 겪지 않도록 만든 실행용 배치 파일입니다.
실행하면 아래 순서로 알아서 준비하고 프로그램을 시작합니다.

1. **Python 감지** — `py` 런처와 `python`을 순서대로 확인합니다. Microsoft 스토어의
   가짜 `python.exe` 별칭(스토어 설치 안내만 띄우는 것)은 실제 실행 테스트로 걸러냅니다.
2. **Python 자동 설치** — Python이 없으면 `winget`으로 Python 3.12를 설치하고,
   winget이 없는 PC에서는 python.org에서 설치 파일을 내려받아 설치합니다.
3. **가상환경 + 패키지 설치** — 같은 폴더에 `.venv` 가상환경을 만들고,
   `requirements.txt`의 패키지(streamlit, pandas, openpyxl, pywin32)를 자동 설치합니다.
4. **프로그램 실행** — `streamlit run autofill_app.py`로 앱을 띄우고 브라우저를 엽니다.

## 설치 방법

1. `WorkReport_AutoFill_실행.bat`과 `requirements.txt`를 `WorkReport_AutoFill_v1.0` 폴더
   (`autofill_app.py`가 있는 곳)에 복사합니다.
2. 앱 파일명이 `autofill_app.py`가 아니라면, 배치 파일을 메모장으로 열어
   상단의 `set "MAIN_SCRIPT=..."` 줄을 실제 파일명으로 바꿉니다.
3. 배치 파일을 더블클릭해 실행합니다. 잠시 후 브라우저에 앱이 열립니다.

## 참고: 수동으로 해결하고 싶다면

- **Python 설치**: https://www.python.org/downloads/ 에서 설치하되,
  첫 화면에서 **"Add python.exe to PATH"를 반드시 체크**하세요.
- **이미 설치했는데도 같은 오류가 나면**: 설정 → 앱 → 고급 앱 설정 → **앱 실행 별칭**에서
  `python.exe`, `python3.exe` 항목을 끄세요.
