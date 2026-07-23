"""
M/H 유효성 검사 모듈
────────────────────────────────────────────────────────────────
핵심 원칙:
  - M/H 합산 기준 = 코드가 분류된 업무 (KPI + Non-KPI) 만
  - 미분류 업무는 실수 가능성이 높으므로 합산에서 제외,
    별도 경고 항목으로 처리
────────────────────────────────────────────────────────────────
"""


def validate(events: list, target_hours: float, threshold: float = 0.5) -> dict:
    """
    이벤트 목록의 M/H 합계를 검증한다.

    Parameters
    ----------
    events       : matched + 포함된 이벤트 리스트
    target_hours : 주간 목표 시간 (예: 40.0)
    threshold    : 허용 오차 (기본 0.5h)

    Returns
    -------
    dict
        classified_hours  : KPI + Non-KPI 합계 (목표 비교 기준)
        total             : 전체 합계 (분류 + 미분류 포함)
        kpi_hours         : KPI 합계
        non_kpi_hours     : Non-KPI 합계
        unclassified_hours: 미분류 합계
        unclassified_count: 미분류 건수
        difference        : classified_hours - target_hours
        is_valid          : abs(difference) <= threshold
        warnings          : 경고 메시지 목록
    """
    included = [e for e in events if e.get("included", True)]

    kpi_h    = round(sum(e["duration_hours"] for e in included
                         if e.get("code_type") == "KPI"), 2)
    non_kpi_h= round(sum(e["duration_hours"] for e in included
                         if e.get("code_type") == "Non-KPI"), 2)
    uncl_h   = round(sum(e["duration_hours"] for e in included
                         if e.get("code_type") == "미분류"), 2)
    uncl_cnt = sum(1 for e in included if e.get("code_type") == "미분류")

    # ★ 목표 비교 기준: 분류된 업무(KPI + Non-KPI)만
    classified = round(kpi_h + non_kpi_h, 2)
    total      = round(classified + uncl_h, 2)     # 참고용 전체 합계

    diff     = round(classified - target_hours, 2)
    is_valid = abs(diff) <= threshold and uncl_cnt == 0   # 미분류가 있으면 valid 아님

    warnings = []

    # M/H 목표 달성 여부 (분류 기준)
    if abs(diff) > threshold:
        direction = "초과" if diff > 0 else "부족"
        warnings.append(
            f"분류 M/H({classified:.1f}h)가 목표({target_hours:.0f}h)보다 "
            f"{abs(diff):.1f}h {direction}합니다."
        )

    # 미분류 경고 — 별도 항목으로 강조
    if uncl_cnt > 0:
        warnings.append(
            f"미분류 {uncl_cnt}건({uncl_h:.1f}h) — Outlook 일정 제목의 코드를 확인하세요."
        )

    return {
        "classified_hours":   classified,     # 목표 비교 기준 (KPI+Non-KPI)
        "total":              total,           # 전체 (분류+미분류, 참고용)
        "target":             target_hours,
        "kpi_hours":          kpi_h,
        "non_kpi_hours":      non_kpi_h,
        "unclassified_hours": uncl_h,
        "unclassified_count": uncl_cnt,
        "difference":         diff,            # classified - target
        "is_valid":           is_valid,        # 분류 기준 충족 + 미분류 없음
        "warnings":           warnings,
    }
