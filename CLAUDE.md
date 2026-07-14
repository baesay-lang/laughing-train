# 업무 포트폴리오 저장소 (Work Portfolio)

이 저장소는 사용자가 Claude와 함께 만든 **최종 결과물**을 모아 GitHub Pages 대시보드로 보여주는 개인 포트폴리오입니다.

- 대시보드 주소: https://baesay-lang.github.io/laughing-train/
- 대시보드는 `main` 브랜치에 푸시될 때 GitHub Actions(`.github/workflows/pages.yml`)가 `scripts/build.py`를 실행해 자동으로 빌드·배포합니다. 수동 배포 작업은 필요 없습니다.

## 사용자가 "포트폴리오에 저장해줘"라고 하면 (모든 Claude 세션 공통 규칙)

1. **최종 결과물만 저장한다.** 중간 산출물(임시 스크립트, 스크래치 노트, 실패한 시도, 로그)은 제외한다. 판단 기준: "한 달 뒤 사용자가 다시 열어볼 가치가 있는 파일인가?"
2. `works/YYYY-MM-DD-짧은-슬러그/` 폴더를 만들고 결과물 파일과 `meta.json`을 넣는다. 날짜는 결과물이 완성된 날짜.
3. `meta.json` 스키마 (title, date, category, description은 필수):

```json
{
  "title": "사람이 읽는 한국어 제목",
  "date": "YYYY-MM-DD",
  "category": "카테고리명",
  "project": "프로젝트명 (같은 주제의 작업을 묶는 이름, 없으면 생략)",
  "description": "무엇을 만들었고 왜 만들었는지 1~3문장",
  "tags": ["선택", "태그"],
  "main": "대표파일.md",
  "external_url": "배포된 결과물이 저장소 밖에 있을 때만 (예: Artifact URL)"
}
```

4. **카테고리는 새로 만들기 전에 기존 것을 재사용한다.** `works/*/meta.json`의 category 값을 먼저 훑고, 의미가 같으면 기존 카테고리를 쓴다. 기본 카테고리: `문서/보고서`, `코드/도구`, `웹/시각화`, `데이터/분석`, `이미지/디자인`, `기타`.
5. **프로젝트 묶기:** 같은 주제·같은 프로젝트의 결과물은 `project` 값을 동일하게 맞춘다. 기존 `project` 값을 먼저 확인해서 표기를 통일한다.
6. **버전 갱신:** 기존 결과물의 새 버전을 저장할 때는 같은 폴더를 재사용한다. 이전 파일은 삭제하지 말고 `archive/v{N}-YYYYMMDD/` 하위 폴더로 옮긴 뒤 새 파일을 넣고, `meta.json`의 `date`를 갱신한다. 대시보드가 archive 폴더를 "이전 버전"으로 자동 표시한다.
7. 커밋 후 `main`에 반영한다(직접 푸시가 가능하면 main에, 아니면 브랜치 + PR). 푸시하면 대시보드는 자동 갱신된다.
8. 로컬에서 미리 확인하려면: `python3 scripts/build.py` 실행 후 `_site/index.html`을 열어본다.

## 저장소 구조

```
works/                  결과물 (폴더 하나 = 결과물 하나)
scripts/build.py        대시보드 생성 스크립트 (Python 표준 라이브러리만 사용)
site/                   대시보드 정적 소스 (HTML/CSS/JS 템플릿)
_site/                  빌드 출력 (gitignore됨, 커밋 금지)
.github/workflows/pages.yml  자동 배포 워크플로우
```

## 주의

- 이 저장소는 **public**이고 대시보드도 공개 URL이다. 민감한 정보(개인정보, 사내 기밀, 자격증명)가 든 파일은 저장하기 전에 사용자에게 확인한다.
- `works/` 폴더명과 `meta.json`만 규칙에 맞으면 나머지는 빌드가 알아서 처리한다. `meta.json`이 없는 폴더도 폴더명(`YYYY-MM-DD-제목`)에서 날짜·제목을 추론해 "미분류"로 표시된다.
