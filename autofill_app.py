# -*- coding: utf-8 -*-
"""
주간업무 자동작성 (WorkReport AutoFill)
────────────────────────────────────────────────────────────────
Outlook 없이 팀 업무코드만으로 월~금 주간업무를 자동 배치하고
기존 주간업무보고 양식 그대로 Excel을 생성한다.
실행:  streamlit run autofill_app.py   (또는 run.bat)
"""
import os
import json
import pandas as pd
import streamlit as st
from datetime import date, timedelta
from pathlib import Path

from src.scheduler import (
    generate_week, copy_last_week, save_history, history_weights,
    day_segments, daily_hours, parse_hhmm,
    load_holidays, holiday_day_events,
)
import re
from src.validator import validate
from src.report_generator_web import generate_report_bytes
from src.outlook_writer import (
    outlook_available, register_events, build_ics, CATEGORY,
)

APP_VERSION = "v1.2.0 (build 2026-07-06)"

BASE       = Path(__file__).parent
CONFIG_DIR = BASE / "config"
TEAMS_DIR  = CONFIG_DIR / "teams"
HIST_DIR   = CONFIG_DIR / "history"
OUTPUT_DIR = BASE / "output"
PREFS_PATH = CONFIG_DIR / "user_prefs.json"

DAY_KO = ["월", "화", "수", "목", "금", "토", "일"]

MODES = {
    "🎲 완전 랜덤":                     "random",
    "🕘 최근 이력 참조 (자주 한 업무 위주)": "history",
    "⭐ 즐겨찾기만 배치":                "favorites",
    "📋 최근 저장 주 그대로 복사":        "copy_last",
}

st.set_page_config(page_title="주간업무 자동작성", page_icon="🗓️",
                   layout="wide", initial_sidebar_state="expanded")


# ── 설정 로드/저장 ────────────────────────────────────────────────────
@st.cache_data
def load_teams() -> list[tuple[str, dict]]:
    teams = []
    for p in sorted(TEAMS_DIR.glob("*.json")):
        if p.stem.startswith("_"):
            continue
        try:
            teams.append((p.stem, json.loads(p.read_text(encoding="utf-8"))))
        except Exception:
            pass
    return teams


def load_prefs() -> dict:
    try:
        return json.loads(PREFS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_prefs(prefs: dict):
    PREFS_PATH.parent.mkdir(parents=True, exist_ok=True)
    PREFS_PATH.write_text(json.dumps(prefs, ensure_ascii=False, indent=1),
                          encoding="utf-8")


def code_label(c: dict) -> str:
    return f"[{c.get('type','?')}] {c.get('level1_name','')} · {c.get('name','')} ({c.get('code','')})"


# ── 세션 초기화 ───────────────────────────────────────────────────────
def _init_state():
    defaults = {
        "this_events": [], "next_events": [],
        "this_nonce": 0,   "next_nonce": 0,
        "prefs": load_prefs(),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _team_pref(team_id: str) -> dict:
    return st.session_state.prefs.setdefault("teams", {}).setdefault(team_id, {})


def _seed_widget(key: str, value):
    if key not in st.session_state:
        st.session_state[key] = value


# ── 이벤트 → DataFrame / 편집 반영 ────────────────────────────────────
def _to_df(events: list) -> pd.DataFrame:
    rows = []
    for ev in events:
        d = ev["date"]
        rows.append({
            "포함": ev.get("included", True),
            "날짜": d.strftime("%m/%d"),
            "요일": DAY_KO[d.weekday()],
            "시작": ev["start"].strftime("%H:%M"),
            "종료": ev["end"].strftime("%H:%M"),
            "M/H":  ev["duration_hours"],
            "업무": ev.get("_label", ""),
            "제목": ev.get("title", ""),
            "유형": ev.get("code_type", "미분류"),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["포함", "날짜", "요일", "시작", "종료", "M/H", "업무", "제목", "유형"])


def _attach_labels(events: list, code_by_code: dict) -> list:
    """이벤트에 selectbox용 라벨을 붙인다 (코드 미존재 시 즉석 라벨)."""
    for ev in events:
        c = code_by_code.get(ev.get("code"))
        ev["_label"] = code_label(c) if c else \
            f"[{ev.get('code_type','?')}] {ev.get('work_name','')} ({ev.get('code','')})"
    return events


def _apply_edits(events: list, edited: pd.DataFrame, by_label: dict) -> list:
    out = []
    for i, ev in enumerate(events):
        if i >= len(edited):
            out.append(ev)
            continue
        row = edited.iloc[i]
        e = dict(ev)
        e["included"] = bool(row["포함"])
        c = by_label.get(str(row["업무"]))
        if c and c["code"] != e.get("code"):
            if e.get("title", "") == e.get("work_name", ""):
                e["title"] = c["name"]
            e["work_name"] = c["name"]
            e["code"]      = c["code"]
            e["code_type"] = c.get("type", "미분류")
            e["_label"]    = code_label(c)
        t = str(row["제목"]).strip()
        if t:
            e["title"] = t
        out.append(e)
    return out


# ── 요약 ──────────────────────────────────────────────────────────────
def _show_summary(events: list, target_h: float):
    if not events:
        st.info("아직 생성된 업무가 없습니다. 상단의 자동 생성 버튼을 누르세요.")
        return
    r = validate(events, target_h)
    tot = r["classified_hours"]
    diff = r["difference"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔵 KPI",     f"{r['kpi_hours']:.1f}h")
    c2.metric("🟢 Non-KPI", f"{r['non_kpi_hours']:.1f}h")
    c3.metric("📊 합계",     f"{tot:.1f}h",
              delta=f"{diff:+.1f}h (목표 {target_h:.0f}h)",
              delta_color="normal" if abs(diff) <= 0.5 else "inverse")
    day_sum: dict = {}
    for e in events:
        if e.get("included", True):
            k = DAY_KO[e["date"].weekday()]
            day_sum[k] = day_sum.get(k, 0.0) + e["duration_hours"]
    c4.metric("📅 일별", " / ".join(f"{day_sum.get(d, 0):.0f}" for d in DAY_KO[:5]),
              delta="월~금 (h)", delta_color="off")


# ── 이벤트 편집기 ─────────────────────────────────────────────────────
def _event_editor(which: str, events: list, labels: list, by_label: dict) -> list:
    if not events:
        return []
    df = _to_df(events)
    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        key=f"editor_{which}_{st.session_state[f'{which}_nonce']}",
        column_config={
            "포함": st.column_config.CheckboxColumn("포함", width=50),
            "날짜": st.column_config.TextColumn("날짜", width=55),
            "요일": st.column_config.TextColumn("요일", width=42),
            "시작": st.column_config.TextColumn("시작", width=55),
            "종료": st.column_config.TextColumn("종료", width=55),
            "M/H":  st.column_config.NumberColumn("M/H", width=52, format="%.1f"),
            "업무": st.column_config.SelectboxColumn("업무 (코드 변경 가능)",
                                                     options=labels, width=430),
            "제목": st.column_config.TextColumn("제목 (보고서 표기)", width=280),
            "유형": st.column_config.TextColumn("유형", width=65),
        },
        disabled=["날짜", "요일", "시작", "종료", "M/H", "유형"],
    )
    new_events = _apply_edits(events, edited, by_label)
    st.session_state[f"{which}_events"] = new_events
    return new_events


# ── 메인 ──────────────────────────────────────────────────────────────
def main():
    _init_state()
    teams = load_teams()
    if not teams:
        st.error("config/teams 에 팀 코드 JSON이 없습니다.")
        st.stop()
    prefs = st.session_state.prefs

    # ── 사이드바 ────────────────────────────────────────────────────
    with st.sidebar:
        st.title("🗓️ 주간업무 자동작성")
        st.caption(f"{APP_VERSION} · 팀 코드 기반 자동 배치")
        st.divider()

        team_names = [cfg.get("team_name", tid) for tid, cfg in teams]
        sel = st.selectbox("🏢 팀(부서) 선택", team_names)
        team_id, team_cfg = teams[team_names.index(sel)]
        codes = team_cfg.get("codes", [])
        by_label = {code_label(c): c for c in codes}
        labels   = list(by_label.keys())
        by_code  = {c["code"]: c for c in codes}
        tp = _team_pref(team_id)

        user_name = st.text_input("👤 성명", value=prefs.get("user_name", ""))
        emp_id    = st.text_input("🔢 사번", value=prefs.get("employee_id", ""))

        st.divider()
        st.subheader("⏰ 근무시간 (부서별 수정 가능)")
        _seed_widget(f"{team_id}_ws", parse_hhmm(tp.get("work_start",  "08:30")))
        _seed_widget(f"{team_id}_ls", parse_hhmm(tp.get("lunch_start", "12:00")))
        _seed_widget(f"{team_id}_le", parse_hhmm(tp.get("lunch_end",   "13:00")))
        _seed_widget(f"{team_id}_we", parse_hhmm(tp.get("work_end",    "17:30")))
        c1, c2 = st.columns(2)
        with c1:
            t_ws = st.time_input("출근",     key=f"{team_id}_ws", step=1800)
            t_ls = st.time_input("점심 시작", key=f"{team_id}_ls", step=1800)
        with c2:
            t_we = st.time_input("퇴근",     key=f"{team_id}_we", step=1800)
            t_le = st.time_input("점심 종료", key=f"{team_id}_le", step=1800)
        segs    = day_segments(t_ws, t_we, t_ls, t_le)
        d_hours = daily_hours(segs)
        target  = d_hours * 5
        st.caption(f"일 {d_hours:.1f}h × 5일 = **주 {target:.1f}h** "
                   f"(점심 {t_ls:%H:%M}~{t_le:%H:%M} 제외)")

        st.divider()
        st.subheader("⚙️ 배치 옵션")
        _seed_widget(f"{team_id}_mode", tp.get("mode", "🎲 완전 랜덤"))
        mode_label = st.radio("배치 방식", list(MODES.keys()),
                              key=f"{team_id}_mode")
        mode = MODES[mode_label]

        _seed_widget(f"{team_id}_min", tp.get("min_block", 1.0))
        _seed_widget(f"{team_id}_max", tp.get("max_block", 4.0))
        c3, c4 = st.columns(2)
        with c3:
            min_h = st.selectbox("업무 최소 시간(h)", [0.5, 1.0, 1.5, 2.0, 3.0, 4.0],
                                 key=f"{team_id}_min")
        with c4:
            max_h = st.selectbox("업무 최대 시간(h)", [2.0, 3.0, 4.0, 4.5, 6.0, 9.0],
                                 key=f"{team_id}_max")
        if max_h < min_h:
            st.warning("최대 시간이 최소보다 작아 최소값으로 맞춥니다.")
            max_h = min_h

        st.divider()
        st.subheader("⭐ 즐겨찾기 업무")
        saved_fav = [c for c in tp.get("favorites", []) if c in by_code]
        _seed_widget(f"{team_id}_fav", [code_label(by_code[c]) for c in saved_fav])
        fav_labels = st.multiselect("자주 쓰는 업무를 등록해 두면 "
                                    "'즐겨찾기만 배치'에서 사용됩니다.",
                                    labels, key=f"{team_id}_fav")
        favorites = [by_label[l]["code"] for l in fav_labels]

        st.divider()
        st.subheader("🏖️ 휴무 처리 코드")
        _auto_hol = next((code_label(c) for c in codes
                          if re.search(r"휴가|휴무|휴일|연차", c.get("name", ""))),
                         labels[0])
        hol_key = f"{team_id}_holcode"
        saved_hol = tp.get("holiday_code")
        _seed_widget(hol_key,
                     code_label(by_code[saved_hol])
                     if saved_hol in by_code else _auto_hol)
        if st.session_state.get(hol_key) not in labels:
            st.session_state[hol_key] = _auto_hol
        hol_label = st.selectbox("공휴일·연차로 지정한 날에 하루 전체로 배치될 코드",
                                 labels, key=hol_key)
        holiday_task = by_label[hol_label]

        st.divider()
        st.subheader("🌏 캘린더 시간대 (ICS/Outlook)")
        TZ_CHOICES = {
            "(UTC+09:00) 서울 — 기본":                       9.0,
            "(UTC+00:00) UTC — 일정이 9시간 밀려 보이면 선택": 0.0,
            "(UTC+08:00) 중국·싱가포르":                      8.0,
            "(UTC+07:00) 베트남·태국":                        7.0,
        }
        _seed_widget("ics_tz", prefs.get("ics_tz", list(TZ_CHOICES)[0]))
        if st.session_state.get("ics_tz") not in TZ_CHOICES:
            st.session_state["ics_tz"] = list(TZ_CHOICES)[0]
        tz_label = st.selectbox(
            "캘린더 계정의 표시 시간대에 맞추세요. "
            "가져온 일정이 9시간 밀려 보인다면 계정이 UTC로 설정된 것이므로 "
            "UTC를 선택하면 보정됩니다.",
            list(TZ_CHOICES), key="ics_tz")
        tz_offset = TZ_CHOICES[tz_label]

        st.divider()
        n_hist = len(list(HIST_DIR.glob(f"{team_id}_*.json")))
        st.caption(f"저장된 이력: {n_hist}주 (Excel 생성 시 자동 저장 → "
                   "'최근 이력 참조'·'복사' 모드에 사용)")

    # 설정 자동 저장
    prefs["user_name"]   = user_name
    prefs["employee_id"] = emp_id
    prefs["ics_tz"]      = tz_label
    tp.update({
        "work_start": t_ws.strftime("%H:%M"), "work_end": t_we.strftime("%H:%M"),
        "lunch_start": t_ls.strftime("%H:%M"), "lunch_end": t_le.strftime("%H:%M"),
        "mode": mode_label, "min_block": min_h, "max_block": max_h,
        "favorites": favorites, "holiday_code": holiday_task["code"],
    })
    save_prefs(prefs)

    # ── 기간 설정 ───────────────────────────────────────────────────
    st.subheader("📅 대상 주 선택")
    today = date.today()
    c1, c2 = st.columns([1, 2])
    with c1:
        base_day = st.date_input("기준일 (해당 주의 월~금 생성)", value=today)
    this_mon = base_day - timedelta(days=base_day.weekday())
    next_mon = this_mon + timedelta(weeks=1)
    with c2:
        st.caption("")
        st.info(f"금주: {this_mon:%m/%d}(월) ~ {this_mon + timedelta(days=4):%m/%d}(금)"
                f"  ·  차주: {next_mon:%m/%d}(월) ~ {next_mon + timedelta(days=4):%m/%d}(금)")

    # ── 휴무일 (공휴일 자동 감지 + 개인 연차 추가) ──────────────────
    hols = load_holidays(CONFIG_DIR,
                         {this_mon.year, (next_mon + timedelta(days=4)).year})
    week_days = ([this_mon + timedelta(days=i) for i in range(5)]
                 + [next_mon + timedelta(days=i) for i in range(5)])

    def _day_label(d: date) -> str:
        base = f"{d:%m/%d}({DAY_KO[d.weekday()]})"
        return f"{base} · {hols[d]}" if d in hols else base

    auto_hols = [d for d in week_days if d in hols]
    sel_hols = st.multiselect(
        f"🏖️ 휴무 처리할 날짜 — 공휴일은 자동 선택됨, 개인 연차는 직접 추가 "
        f"(해당 일은 '{holiday_task['name']}' 코드로 하루 전체 배치)",
        week_days, default=auto_hols, format_func=_day_label,
        key=f"hols_{this_mon.isoformat()}",
    )
    hol_map_all = {d: hols.get(d, "개인 연차") for d in sel_hols}

    # ── 생성 함수 ───────────────────────────────────────────────────
    def _generate(which: str, monday: date):
        week_hols = {d: n for d, n in hol_map_all.items()
                     if monday <= d <= monday + timedelta(days=4)}
        wk_kwargs = dict(
            work_start=t_ws.strftime("%H:%M"), work_end=t_we.strftime("%H:%M"),
            lunch_start=t_ls.strftime("%H:%M"), lunch_end=t_le.strftime("%H:%M"))
        try:
            if mode == "copy_last":
                ev = copy_last_week(HIST_DIR, team_id, monday,
                                    exclude_monday=monday)
                if ev is None:
                    st.error("복사할 저장 이력이 없습니다. 먼저 한 주를 생성하고 "
                             "Excel을 만들어 이력을 남기세요.")
                    return
                # 복사본에도 휴무일 적용: 해당 일 업무를 휴무 코드로 교체
                if week_hols:
                    ev = [e for e in ev if e["date"] not in week_hols]
                    for d, name in week_hols.items():
                        ev.extend(holiday_day_events(d, holiday_task, name,
                                                     **wk_kwargs))
                    ev.sort(key=lambda e: e["start"])
            else:
                hw = history_weights(HIST_DIR, team_id) if mode == "history" else {}
                if mode == "history" and not hw:
                    st.warning("저장된 이력이 없어 이번에는 전체 코드에서 랜덤 배치합니다.")
                if mode == "favorites" and not favorites:
                    st.error("즐겨찾기가 비어 있습니다. 사이드바에서 먼저 등록하세요.")
                    return
                ev = generate_week(
                    monday, codes, mode=mode, favorites=favorites,
                    history_weights=hw, min_h=min_h, max_h=max_h,
                    holiday_map=week_hols, holiday_task=holiday_task,
                    **wk_kwargs,
                )
            _attach_labels(ev, by_code)
            st.session_state[f"{which}_events"] = ev
            st.session_state[f"{which}_nonce"] += 1
            st.toast(f"{'금주' if which == 'this' else '차주'} {len(ev)}건 배치 완료", icon="✅")
        except Exception as e:
            st.error(f"생성 오류: {e}")

    b1, b2, b3 = st.columns(3)
    with b1:
        if st.button("▶️  금주 자동 생성", use_container_width=True, type="primary"):
            _generate("this", this_mon)
    with b2:
        if st.button("▶️  차주 자동 생성", use_container_width=True, type="primary"):
            _generate("next", next_mon)
    with b3:
        if st.button("⏩  금주+차주 한번에 생성", use_container_width=True):
            _generate("this", this_mon)
            _generate("next", next_mon)

    st.divider()

    # ── 금주 / 차주 탭 ──────────────────────────────────────────────
    tab1, tab2 = st.tabs(["📅 금주 (ACTUAL)", "📅 차주 (PLAN)"])
    with tab1:
        this_events = _event_editor("this",
                                    _attach_labels(st.session_state.this_events, by_code),
                                    labels, by_label)
        _show_summary(this_events, target)
    with tab2:
        next_events = _event_editor("next",
                                    _attach_labels(st.session_state.next_events, by_code),
                                    labels, by_label)
        _show_summary(next_events, target)

    # ── Outlook 자동 기입 ───────────────────────────────────────────
    st.divider()
    st.subheader("📆 Outlook 캘린더 자동 기입")

    if "outlook_status" not in st.session_state:
        st.session_state.outlook_status = outlook_available()
    ol_ok, ol_msg = st.session_state.outlook_status

    if ol_ok:
        st.caption(
            f"연결됨: {ol_msg} · 제목은 `코드 [KPI] 업무명` 형식으로 등록되어 "
            "기존 주간업무보고(아웃룩 추출)와 호환됩니다. "
            f"자동 등록 일정에는 범주 '{CATEGORY}'가 붙습니다."
        )
        replace_flag = st.checkbox(
            f"등록 전 같은 주의 기존 '{CATEGORY}' 일정 삭제 후 교체 (중복 방지)",
            value=True,
        )
        oc1, oc2 = st.columns(2)
        with oc1:
            if st.button("📆  금주 일정 Outlook 등록", use_container_width=True):
                evs = [e for e in this_events if e.get("included", True)]
                if not evs:
                    st.error("먼저 금주 업무를 생성하세요.")
                else:
                    try:
                        with st.spinner("Outlook에 등록 중..."):
                            n, d = register_events(evs, this_mon, replace=replace_flag)
                        st.success(f"금주 {n}건 등록 완료"
                                   + (f" (기존 자동등록 {d}건 교체)" if d else ""))
                    except Exception as e:
                        st.error(f"Outlook 등록 오류: {e}")
        with oc2:
            if st.button("📆  차주 일정 Outlook 등록", use_container_width=True):
                evs = [e for e in next_events if e.get("included", True)]
                if not evs:
                    st.error("먼저 차주 업무를 생성하세요.")
                else:
                    try:
                        with st.spinner("Outlook에 등록 중..."):
                            n, d = register_events(evs, next_mon, replace=replace_flag)
                        st.success(f"차주 {n}건 등록 완료"
                                   + (f" (기존 자동등록 {d}건 교체)" if d else ""))
                    except Exception as e:
                        st.error(f"Outlook 등록 오류: {e}")
    else:
        st.info(f"이 PC에서 Outlook 직접 등록을 쓸 수 없습니다 ({ol_msg}). "
                "아래 ICS 파일을 내려받아 Outlook에서 열면 일정이 추가됩니다.")

    with st.expander("📎 ICS 파일로 내려받기 (Outlook 없이 가져오기)"):
        ic1, ic2 = st.columns(2)
        with ic1:
            if this_events:
                st.download_button(
                    "금주 ICS 다운로드", data=build_ics(this_events, tz_offset),
                    file_name=f"주간업무_금주_{this_mon:%Y%m%d}.ics",
                    mime="text/calendar", use_container_width=True)
        with ic2:
            if next_events:
                st.download_button(
                    "차주 ICS 다운로드", data=build_ics(next_events, tz_offset),
                    file_name=f"주간업무_차주_{next_mon:%Y%m%d}.ics",
                    mime="text/calendar", use_container_width=True)

    # ── Excel 생성 ──────────────────────────────────────────────────
    st.divider()
    col_gen, col_dl = st.columns([1, 2])
    with col_gen:
        gen = st.button("💾  Excel 보고서 생성", type="primary",
                        use_container_width=True)

    if gen:
        if not this_events and not next_events:
            st.error("먼저 금주 또는 차주 업무를 생성하세요.")
        else:
            try:
                sheet = f"{this_mon:%y%m%d}_{next_mon + timedelta(days=4):%y%m%d}"
                xl = generate_report_bytes(
                    [e for e in this_events if e.get("included", True)],
                    [e for e in next_events if e.get("included", True)],
                    user_name or "미입력", emp_id or "미입력", sheet,
                )
                team_name = team_cfg.get("team_name", "팀")
                base = f"주간업무보고_{team_name}_{this_mon:%Y%m%d}"
                ver = 1
                while (OUTPUT_DIR / f"{base}_v{ver}.xlsx").exists():
                    ver += 1
                fname = f"{base}_v{ver}.xlsx"
                OUTPUT_DIR.mkdir(exist_ok=True)
                (OUTPUT_DIR / fname).write_bytes(xl)

                # 이력 저장 (최근 이력 참조 / 복사 모드의 재료)
                if this_events:
                    save_history(HIST_DIR, team_id, this_mon, this_events)
                if next_events:
                    save_history(HIST_DIR, team_id, next_mon, next_events)

                with col_dl:
                    st.download_button("📥  Excel 다운로드", data=xl,
                                       file_name=fname,
                                       mime="application/vnd.openxmlformats-"
                                            "officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                st.success(f"생성 완료: output\\{fname} (이력 저장됨)")
                try:
                    os.startfile(OUTPUT_DIR)          # 탐색기로 열기
                except Exception:
                    pass
            except Exception as e:
                st.error(f"Excel 생성 오류: {e}")


if __name__ == "__main__":
    main()
