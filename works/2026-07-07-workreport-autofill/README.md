# 주간업무 자동작성 (WorkReport AutoFill)

팀 업무코드만으로 월~금 주간업무를 자동 배치하고, 기존 주간업무보고 양식
그대로 Excel을 생성하는 Streamlit 웹앱.

## 주요 기능

- **자동 배치**: 완전 랜덤 / 최근 이력 참조 / 즐겨찾기만 / 최근 주 복사 4가지 모드
- **공휴일·연차 처리**: 한국 공휴일 자동 감지(2026~2028), 지정일은 휴가 코드로 하루 전체 배치
- **표 편집**: 생성 결과를 표에서 직접 수정 (업무 코드 변경, 제목 수정, 포함/제외)
- **Excel 보고서**: 기존 주간업무보고 양식(KPI/Non-KPI 분류, 비중, DataBar) 그대로 생성
- **Outlook 연동**: 로컬 Windows에선 캘린더 직접 등록, 웹에서는 ICS 파일 다운로드

## 웹으로 사용 (회사 PC — 설치 불가 환경)

[Streamlit Community Cloud](https://share.streamlit.io)에 이 저장소를 연결해 배포:

1. share.streamlit.io 접속 → GitHub 계정으로 로그인
2. **Create app** → 이 저장소 / 배포할 브랜치 선택
3. **Main file path**: `autofill_app.py` → **Deploy**
4. 발급된 `https://<앱이름>.streamlit.app` URL로 어느 PC에서든 접속

> ⚠️ 웹 배포 시 참고
> - Outlook 직접 등록은 서버에서 불가 → **ICS 다운로드**로 대체 (내려받아 열면 Outlook에 추가됨)
> - 서버 저장 데이터(즐겨찾기 설정, 주간 이력)는 앱 재배포/재시작 시 초기화될 수 있음
> - 같은 URL을 여러 명이 쓰면 설정·이력이 공유됨 (개인별 사용은 각자 배포 권장)

## 로컬 실행 (Windows)

`WorkReport_AutoFill_실행.bat` 더블클릭 — Python이 없으면 자동 설치하고,
가상환경·패키지 설치 후 브라우저를 엽니다. 자세한 내용은
[LAUNCHER_사용법.md](LAUNCHER_사용법.md) 참조.

수동 실행:

```bash
pip install -r requirements.txt
streamlit run autofill_app.py
```

## 구성

```
autofill_app.py              # Streamlit 메인 앱
src/scheduler.py             # 주간 자동 배치 엔진 (블록 분할·모드별 선택·이력)
src/validator.py             # M/H 합계 검증 (KPI/Non-KPI/미분류)
src/report_generator_web.py  # Excel 보고서 생성 (openpyxl)
src/outlook_writer.py        # Outlook 등록(win32com) + ICS 생성
config/holidays_kr.json      # 한국 공휴일 (2026~2028)
config/teams/*.json          # 팀별 업무코드 정의
```

## 팀 추가 방법

`config/teams/새팀.json` 파일을 추가하면 앱의 팀 선택 목록에 자동 표시:

```json
{
  "team_name": "팀 이름",
  "department": "본부명",
  "codes": [
    {"code": "GA07-12", "type": "Non-KPI", "level1": "GA07",
     "level1_name": "General", "level2": "12", "name": "교육/훈련"}
  ]
}
```
