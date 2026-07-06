# -*- coding: utf-8 -*-
"""
Outlook 캘린더 자동 기입 모듈 (데스크톱 Outlook, win32com)
────────────────────────────────────────────────────────────────
- 일정 제목: "{코드} [{KPI/Non-KPI}] {업무명}"
  → 기존 주간업무보고(WorkReport)의 CodeMatcher 가 그대로 인식하는 형식
- 자동 등록 일정에는 고유 카테고리를 붙여, 재등록 시 그 카테고리만
  안전하게 삭제·교체한다 (사용자가 직접 만든 일정은 건드리지 않음).
- Outlook 이 없는 PC를 위해 .ics 파일 생성 함수도 제공.
"""
from datetime import date, datetime, timedelta

CATEGORY = "주간업무자동"          # 자동 등록 표식 (Outlook 범주)
OL_APPOINTMENT = 1                # olAppointmentItem
OL_CALENDAR = 9                   # olFolderCalendar
OL_BUSY = 2                       # olBusy


def subject_of(ev: dict) -> str:
    code = ev.get("code", "")
    ctype = ev.get("code_type", "")
    title = ev.get("title", ev.get("work_name", ""))
    if code and code != "미분류":
        return f"{code} [{ctype}] {title}"
    return title


def _com_init():
    """Streamlit 스크립트 스레드에서 COM 사용을 위한 초기화."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _has_outlook_profile() -> bool:
    """Outlook 계정(프로필) 존재 여부 — 미설정 PC에서 시작 마법사에
    걸려 멈추는 것을 막기 위한 사전 확인."""
    try:
        import winreg
    except ImportError:
        return True                      # 확인 불가 시 시도는 해 본다
    for ver in ("16.0", "15.0", "14.0"):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{ver}\Outlook\Profiles")
            if winreg.QueryInfoKey(key)[0] > 0:
                return True
        except OSError:
            continue
    return False


def outlook_available(timeout: float = 6.0) -> tuple[bool, str]:
    """데스크톱 Outlook COM 사용 가능 여부 (앱이 멈추지 않도록
    프로필 사전 확인 + 타임아웃 적용)."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "pywin32 미설치 (pip install pywin32)"

    if not _has_outlook_profile():
        return False, "Outlook 계정(프로필) 미설정"

    import threading
    result: dict = {}

    def _probe():
        try:
            _com_init()
            import win32com.client
            app = win32com.client.Dispatch("Outlook.Application")
            result["ver"] = app.Version
        except Exception as e:
            result["err"] = str(e)

    t = threading.Thread(target=_probe, daemon=True)
    t.start()
    t.join(timeout)
    if t.is_alive():
        return False, "Outlook 응답 없음 (초기 설정 창 대기 중일 수 있음)"
    if "ver" in result:
        return True, f"Outlook {result['ver']}"
    return False, f"Outlook 실행 불가: {result.get('err', '알 수 없음')}"


def _week_range(monday: date) -> tuple[datetime, datetime]:
    """삭제 검색 범위. 과거 시간대 버그로 ±9h 밀려 등록된 일정까지
    쓸어내도록 일요일 00:00 ~ 토요일 24:00 로 하루씩 넓게 잡는다.
    (카테고리가 붙은 자동 등록 일정만 지우므로 안전)"""
    start = datetime.combine(monday, datetime.min.time()) - timedelta(days=1)
    end = start + timedelta(days=7)      # 일 00:00 ~ 일 00:00
    return start, end


def _aware(dt: datetime) -> datetime:
    """naive datetime → 로컬 시간대 명시 (COM 변환 시 UTC 오해석 방지)."""
    if dt.tzinfo is None:
        return dt.astimezone()
    return dt


def delete_auto_events(monday: date) -> int:
    """해당 주(월~금)에서 CATEGORY 가 붙은 자동 등록 일정만 삭제."""
    _com_init()
    import win32com.client
    app = win32com.client.Dispatch("Outlook.Application")
    cal = app.GetNamespace("MAPI").GetDefaultFolder(OL_CALENDAR)
    start, end = _week_range(monday)

    items = cal.Items
    items.IncludeRecurrences = False
    items.Sort("[Start]")
    # Outlook 날짜 필터는 en-US 형식 필수 (기존 outlook_reader와 동일)
    flt = (f"[Start] >= '{start.strftime('%m/%d/%Y 12:00 AM')}' AND "
           f"[Start] < '{end.strftime('%m/%d/%Y 12:00 AM')}'")
    targets = []
    for it in items.Restrict(flt):
        try:
            cats = (it.Categories or "")
            if CATEGORY in [c.strip() for c in cats.split(",")]:
                targets.append(it)
        except Exception:
            continue
    n = 0
    for it in reversed(targets):
        try:
            it.Delete()
            n += 1
        except Exception:
            continue
    return n


def register_events(events: list[dict], monday: date,
                    replace: bool = True) -> tuple[int, int]:
    """
    이벤트를 Outlook 기본 캘린더에 등록.
    replace=True 면 같은 주의 기존 자동 등록 일정을 먼저 삭제.
    Returns (등록 건수, 삭제 건수)
    """
    _com_init()
    import win32com.client
    deleted = delete_auto_events(monday) if replace else 0

    app = win32com.client.Dispatch("Outlook.Application")
    n = 0
    for ev in events:
        if not ev.get("included", True):
            continue
        appt = app.CreateItem(OL_APPOINTMENT)
        appt.Subject = subject_of(ev)
        appt.Start = _aware(ev["start"])  # 시간대 명시 → 9시간 밀림 방지
        appt.End = _aware(ev["end"])
        appt.Categories = CATEGORY
        appt.BusyStatus = OL_BUSY
        appt.ReminderSet = False
        appt.Body = f"주간업무 자동작성 등록 ({ev.get('code_type','')} / {ev.get('code','')})"
        appt.Save()
        n += 1
    return n, deleted


# ── ICS (Outlook COM 이 없을 때 가져오기용) ──────────────────────────
def build_ics(events: list[dict]) -> str:
    def esc(s: str) -> str:
        return (s.replace("\\", "\\\\").replace(";", r"\;")
                 .replace(",", r"\,").replace("\n", r"\n"))

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//WorkReport AutoFill//KR",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    for i, ev in enumerate(events):
        if not ev.get("included", True):
            continue
        uid = f"wra-{ev['start']:%Y%m%dT%H%M}-{i}@autofill"
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{ev['start']:%Y%m%dT%H%M%S}",
            f"DTSTART:{ev['start']:%Y%m%dT%H%M%S}",
            f"DTEND:{ev['end']:%Y%m%dT%H%M%S}",
            f"SUMMARY:{esc(subject_of(ev))}",
            f"CATEGORIES:{CATEGORY}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"
