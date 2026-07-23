# -*- coding: utf-8 -*-
"""
주간업무 자동 배치 엔진
────────────────────────────────────────────────────────────────
- 근무시간(출근~퇴근, 점심 제외)을 0.5h 단위 블록으로 쪼개고
  팀 업무코드를 배치해 월~금 이벤트 목록을 생성한다.
- 배치 방식: random(완전 랜덤) / history(최근 이력 가중) /
  favorites(즐겨찾기만) / copy_last(최근 저장 주 복사)
- 이벤트 형식은 기존 주간업무보고(report_generator_web)와 동일.
"""
import json
import random
from datetime import datetime, date, time, timedelta
from pathlib import Path

STEP = 0.5  # 최소 시간 단위 (30분)


# ── 시간 유틸 ─────────────────────────────────────────────────────────
def parse_hhmm(s: str) -> time:
    h, m = s.split(":")
    return time(int(h), int(m))


def day_segments(work_start: time, work_end: time,
                 lunch_start: time, lunch_end: time) -> list[tuple[time, time]]:
    """점심시간을 제외한 근무 구간 목록. 점심이 근무시간 밖이면 통구간."""
    segs = []
    if lunch_start <= work_start or lunch_end >= work_end or lunch_start >= lunch_end:
        segs.append((work_start, work_end))
    else:
        segs.append((work_start, lunch_start))
        segs.append((lunch_end, work_end))
    return [(a, b) for a, b in segs if a < b]


def _hours_between(a: time, b: time) -> float:
    return ((b.hour * 60 + b.minute) - (a.hour * 60 + a.minute)) / 60.0


def daily_hours(segments: list[tuple[time, time]]) -> float:
    return round(sum(_hours_between(a, b) for a, b in segments), 2)


# ── 블록 분할 ─────────────────────────────────────────────────────────
def split_hours(total: float, min_h: float, max_h: float,
                rng: random.Random) -> list[float]:
    """
    total 시간을 0.5h 배수 블록들로 분할.
    각 블록은 min_h 이상, max_h 이하 (분할 불가능하면 통블록).
    """
    total = round(total * 2) / 2
    min_h = max(STEP, round(min_h * 2) / 2)
    max_h = max(min_h, round(max_h * 2) / 2)
    sizes: list[float] = []
    rem = total
    while rem > 1e-9:
        if rem < min_h * 2:            # 더 쪼개면 min_h 미달 → 통으로
            sizes.append(round(rem, 2))
            break
        cands = []
        s = min_h
        while s <= min(max_h, rem) + 1e-9:
            r2 = round(rem - s, 2)
            if r2 < 1e-9 or r2 >= min_h - 1e-9:
                cands.append(round(s, 2))
            s = round(s + STEP, 2)
        if not cands:
            sizes.append(round(rem, 2))
            break
        pick = rng.choice(cands)
        sizes.append(pick)
        rem = round(rem - pick, 2)
    return sizes


# ── 업무 선택 ─────────────────────────────────────────────────────────
def _pick(pool: list[dict], weights: dict, rng: random.Random,
          avoid: str | None) -> dict:
    """가중 랜덤 선택. 직전 블록과 같은 코드는 가능하면 회피."""
    items = [c for c in pool if c["code"] != avoid] or pool
    w = [max(0.01, float(weights.get(c["code"], 1.0))) for c in items]
    return rng.choices(items, weights=w, k=1)[0]


def build_pool(codes: list[dict], mode: str, favorites: list[str],
               history_weights: dict) -> tuple[list[dict], dict]:
    """모드별 (선택 풀, 가중치) 구성."""
    if mode == "favorites":
        pool = [c for c in codes if c["code"] in set(favorites)]
        return pool, {}
    if mode == "history":
        seen = {k: v for k, v in history_weights.items() if v > 0}
        pool = [c for c in codes if c["code"] in seen]
        if pool:
            return pool, seen
        return list(codes), {}          # 이력 없으면 전체 랜덤
    return list(codes), {}              # random


# ── 공휴일 ────────────────────────────────────────────────────────────
def load_holidays(config_dir: Path, years: set[int]) -> dict:
    """한국 공휴일 {date: 이름}. holidays 패키지 우선, 없으면
    config/holidays_kr.json 사용. holidays_extra.json(회사 지정 휴무일)은
    항상 병합한다."""
    out: dict = {}
    p = Path(config_dir) / "holidays_kr.json"
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
        out.update({date.fromisoformat(k): v for k, v in raw.items()})
    except Exception:
        pass
    try:
        import holidays as _h
        for d, name in _h.KR(years=sorted(years)).items():
            out[d] = name
    except Exception:
        pass
    try:
        extra = json.loads((Path(config_dir) / "holidays_extra.json")
                           .read_text(encoding="utf-8"))
        out.update({date.fromisoformat(k): v for k, v in extra.items()})
    except Exception:
        pass
    return out


def holiday_day_events(d: date, task: dict, holiday_name: str = "", *,
                       work_start: str = "08:30", work_end: str = "17:30",
                       lunch_start: str = "12:00",
                       lunch_end: str = "13:00") -> list[dict]:
    """하루 전체를 휴가/휴무 코드로 채운 이벤트 목록 (구간당 1블록)."""
    segs = day_segments(parse_hhmm(work_start), parse_hhmm(work_end),
                        parse_hhmm(lunch_start), parse_hhmm(lunch_end))
    title = f"{task['name']} ({holiday_name})" if holiday_name else task["name"]
    evs = []
    for a, b in segs:
        evs.append({
            "date":           d,
            "start":          datetime.combine(d, a),
            "end":            datetime.combine(d, b),
            "duration_hours": _hours_between(a, b),
            "title":          title,
            "work_name":      task["name"],
            "code":           task["code"],
            "code_type":      task.get("type", "Non-KPI"),
            "included":       True,
        })
    return evs


# ── 주간 생성 ─────────────────────────────────────────────────────────
def generate_week(monday: date, codes: list[dict], *,
                  mode: str = "random",
                  favorites: list[str] | None = None,
                  history_weights: dict | None = None,
                  min_h: float = 1.0,
                  max_h: float = 4.0,
                  work_start: str = "08:30", work_end: str = "17:30",
                  lunch_start: str = "12:00", lunch_end: str = "13:00",
                  holiday_map: dict | None = None,
                  holiday_task: dict | None = None,
                  seed: int | None = None) -> list[dict]:
    """월~금 5일 자동 배치 이벤트 목록 생성.
    holiday_map({date: 이름})에 든 날은 holiday_task 코드로 하루 전체 배치."""
    rng = random.Random(seed)
    favorites = favorites or []
    history_weights = history_weights or {}
    holiday_map = holiday_map or {}

    pool, weights = build_pool(codes, mode, favorites, history_weights)
    if not pool:
        raise ValueError("배치할 업무가 없습니다. 즐겨찾기 또는 팀 코드를 확인하세요.")

    segs = day_segments(parse_hhmm(work_start), parse_hhmm(work_end),
                        parse_hhmm(lunch_start), parse_hhmm(lunch_end))
    events = []
    for wd in range(5):                              # 월~금
        d = monday + timedelta(days=wd)
        if d in holiday_map and holiday_task:
            events.extend(holiday_day_events(
                d, holiday_task, holiday_map[d],
                work_start=work_start, work_end=work_end,
                lunch_start=lunch_start, lunch_end=lunch_end))
            continue
        prev_code = None
        for seg_start, seg_end in segs:
            cursor = datetime.combine(d, seg_start)
            for size in split_hours(_hours_between(seg_start, seg_end),
                                    min_h, max_h, rng):
                task = _pick(pool, weights, rng, avoid=prev_code)
                prev_code = task["code"]
                st_dt = cursor
                en_dt = cursor + timedelta(hours=size)
                cursor = en_dt
                events.append({
                    "date":           d,
                    "start":          st_dt,
                    "end":            en_dt,
                    "duration_hours": size,
                    "title":          task["name"],
                    "work_name":      task["name"],
                    "code":           task["code"],
                    "code_type":      task.get("type", "미분류"),
                    "included":       True,
                })
    return events


# ── 이력 저장/로드 ────────────────────────────────────────────────────
def _hist_path(hist_dir: Path, team_id: str, monday: date) -> Path:
    return hist_dir / f"{team_id}_{monday.isoformat()}.json"


def save_history(hist_dir: Path, team_id: str, monday: date,
                 events: list[dict]) -> Path:
    hist_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    for e in events:
        if not e.get("included", True):
            continue
        rows.append({
            "date":           e["date"].isoformat(),
            "start":          e["start"].isoformat(),
            "end":            e["end"].isoformat(),
            "duration_hours": e["duration_hours"],
            "title":          e.get("title", ""),
            "work_name":      e.get("work_name", ""),
            "code":           e.get("code", ""),
            "code_type":      e.get("code_type", "미분류"),
        })
    p = _hist_path(hist_dir, team_id, monday)
    p.write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    return p


def _load_events(p: Path) -> list[dict]:
    rows = json.loads(p.read_text(encoding="utf-8"))
    out = []
    for r in rows:
        out.append({
            "date":           date.fromisoformat(r["date"]),
            "start":          datetime.fromisoformat(r["start"]),
            "end":            datetime.fromisoformat(r["end"]),
            "duration_hours": r["duration_hours"],
            "title":          r.get("title", ""),
            "work_name":      r.get("work_name", ""),
            "code":           r.get("code", ""),
            "code_type":      r.get("code_type", "미분류"),
            "included":       True,
        })
    return out


def list_history(hist_dir: Path, team_id: str) -> list[Path]:
    """해당 팀 이력 파일을 주차 오름차순으로."""
    if not hist_dir.exists():
        return []
    return sorted(hist_dir.glob(f"{team_id}_*.json"))


def history_weights(hist_dir: Path, team_id: str, recent_n: int = 8) -> dict:
    """최근 n개 주차의 코드별 누적 시간 → 가중치."""
    w: dict = {}
    for p in list_history(hist_dir, team_id)[-recent_n:]:
        try:
            for e in _load_events(p):
                w[e["code"]] = w.get(e["code"], 0.0) + float(e["duration_hours"])
        except Exception:
            continue
    return w


def copy_last_week(hist_dir: Path, team_id: str, target_monday: date,
                   exclude_monday: date | None = None) -> list[dict] | None:
    """가장 최근 저장 주를 대상 주로 날짜만 이동해 복사."""
    files = list_history(hist_dir, team_id)
    if exclude_monday:
        files = [p for p in files
                 if p.stem != f"{team_id}_{exclude_monday.isoformat()}"]
    if not files:
        return None
    src = _load_events(files[-1])
    if not src:
        return None
    base_monday = min(e["date"] for e in src)
    base_monday -= timedelta(days=base_monday.weekday())
    out = []
    for e in src:
        shift = (e["date"] - base_monday).days
        if shift > 4:                     # 주말 데이터는 제외
            continue
        nd = target_monday + timedelta(days=shift)
        out.append({**e,
                    "date":  nd,
                    "start": datetime.combine(nd, e["start"].time()),
                    "end":   datetime.combine(nd, e["end"].time()),
                    "included": True})
    return out or None
