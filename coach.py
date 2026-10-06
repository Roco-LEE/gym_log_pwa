# -*- coding: utf-8 -*-
"""LLM 운동 코치 (1단계) — 엑셀 기록을 읽어 다음 운동 계획을 짜고 plan.json 에 넣는다.
   설계: docs/LLM코치_1단계_설계.md

   python coach.py --date 2026-10-08 --focus 하체 --note "허리 약간 뻐근" --context
       → LLM 없이 컨텍스트(LLM 에 보낼 JSON)만 출력. 통계가 맞는지 눈으로 확인용
   python coach.py --date 2026-10-08 --focus 하체 --note "허리 약간 뻐근" --dry-run
       → LLM 으로 계획을 받아 출력만 (plan.json 은 안 씀)

- 통계(지난 세트·증량 신호·e1RM·부위별 마지막 날짜)는 파이썬이 계산해서 준다. LLM 에게 계산을 시키지 않는다.
- 개인 설정(목표·부상·주당 횟수·시간)은 coach_profile.json (git 제외). 없으면 기본값.
- LLM 출력은 구조화 출력(output_config.format = JSON 스키마)으로 받는다. 운동 이름은 스키마 enum.
  (설계의 '도구 강제 호출'은 Sonnet 5.5·Opus 5.5 에서 400 이라, 모델을 바꿔도 그대로 되는 쪽으로)
- API 키는 OS 환경변수 ANTHROPIC_API_KEY 에만. 이 폴더에 두지 않는다 (serve.py 가 폴더를 서빙함).
"""
import argparse, datetime, json, os, re, sys, time
import sync_excel

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.path.join(HERE, "coach_profile.json")
MODEL = "claude-haiku-4-5"   # 저가로 시작, 계획이 아쉬우면 --model claude-sonnet-5-5
MAX_TOKENS = 8000
WINDOW_DAYS = 42          # 최근 범위: 6주 또는 최근 12세션 중 짧은 쪽
WINDOW_SESSIONS = 12
MEMOS_PER_EX = 3          # 운동별로 보여줄 최근 메모 수
DAY_TAG = "[오늘] "
DEFAULT_PROFILE = {"goal": "근비대·자세 우선", "injuries": [], "days_per_week": 3, "time_budget_min": 60}


def load_profile():
    if not os.path.exists(PROFILE):
        print(f"※ {os.path.basename(PROFILE)} 없음 → 기본 설정 사용", file=sys.stderr)
        return dict(DEFAULT_PROFILE)
    with open(PROFILE, encoding="utf-8") as f:
        return {**DEFAULT_PROFILE, **json.load(f)}


def rep_range(text):
    """'10-15' → (10, 15), '30-60초' → (30, 60), '12' → (12, 12), '' → None"""
    n = [int(x) for x in re.findall(r"\d+", str(text or ""))]
    return (n[0], n[-1]) if n else None


def e1rm(kg, reps):
    """Epley. 맨몸(0kg)·기록 없음은 None"""
    return round(kg * (1 + reps / 30), 1) if kg and reps else None


def split_day_memo(memo):
    """[기록] W열 '[오늘] 그날 메모 · 운동 메모' → ('그날 메모', '운동 메모')"""
    memo = (memo or "").strip()
    if not memo.startswith(DAY_TAG):
        return "", memo
    day, _, rest = memo[len(DAY_TAG):].partition(" · ")
    return day.strip(), rest.strip()


def build_context(data, target_date, focus=None, note="", time_budget_min=None, profile=None):
    """sync_excel.read_excel() 결과 → LLM 에 보낼 컨텍스트(dict). target_date: 'YYYY-MM-DD'"""
    profile = profile or dict(DEFAULT_PROFILE)
    tgt = datetime.date.fromisoformat(target_date)
    ex_info = {e["n"]: e for e in data["exercises"]}
    # 부위 없는 항목(트레드밀 등 유산소)은 계획 후보에서 뺌
    exercises = [{"name": e["n"], "part": e["p"], "type": e["t"], "step": e["step"], "rep_range": e["rep"]}
                 for e in data["exercises"] if e["p"]]

    past = sorted((r for r in data["records"] if r["date"] < target_date), key=lambda r: r["date"])
    done_today = [r for r in data["records"] if r["date"] == target_date]

    # 최근 범위 시작일: 6주 전과 12번째 최근 세션 날짜 중 늦은 쪽
    dates = sorted({r["date"] for r in past}, reverse=True)
    since = (tgt - datetime.timedelta(days=WINDOW_DAYS)).isoformat()
    if len(dates) >= WINDOW_SESSIONS:
        since = max(since, dates[WINDOW_SESSIONS - 1])
    recent = [r for r in past if r["date"] >= since]

    # 운동별 통계 — 지난 세트는 기간과 상관없이 마지막 기록(증량 상한 계산에 필요), 횟수·메모는 최근 범위만
    by_ex = {}
    for r in past:
        by_ex.setdefault(r["ex"], []).append(r)
    stats = []
    for ex, rows in by_ex.items():
        last = rows[-1]
        info = ex_info.get(ex, {})
        rng = rep_range(info.get("rep"))
        reps = [s["reps"] for s in last["sets"]]
        memos = []
        for r in reversed(rows):
            m = split_day_memo(r["memo"])[1]
            if m and r["date"] >= since and len(memos) < MEMOS_PER_EX:
                memos.append({"date": r["date"], "memo": m})
        e1 = [e1rm(s["kg"], s["reps"]) for r in rows for s in r["sets"]]
        e1 = [x for x in e1 if x]
        stats.append({
            "ex": ex,
            "part": info.get("p", ""),
            "last_date": last["date"],
            "days_since": (tgt - datetime.date.fromisoformat(last["date"])).days,
            "last_warm": last["warm"],
            "last_sets": last["sets"],
            "last_top_kg": max((s["kg"] for s in last["sets"]), default=None),
            "hit_top_of_range": bool(rng and reps and all(x >= rng[1] for x in reps)),
            "best_e1rm": max(e1) if e1 else None,
            "sessions_recent": sum(1 for r in rows if r["date"] >= since),
            "memos": memos,
        })
    stats.sort(key=lambda s: s["last_date"], reverse=True)

    part_last = {}
    for r in past:
        p = ex_info.get(r["ex"], {}).get("p")
        if p:
            part_last[p] = r["date"]          # past 가 날짜순이라 마지막 값이 최신

    day_memos = {}
    for r in recent:
        d = split_day_memo(r["memo"])[0]
        if d:
            day_memos[r["date"]] = d

    return {
        "target_date": target_date,
        "request": {"focus": focus, "note": note or "", "time_budget_min": time_budget_min or profile["time_budget_min"]},
        "profile": {k: profile[k] for k in ("goal", "injuries", "days_per_week")},
        "window": {"since": since, "sessions": len({r["date"] for r in recent})},
        "exercises": exercises,
        "stats": stats,
        "part_last_trained": dict(sorted(part_last.items(), key=lambda kv: kv[1], reverse=True)),
        "recent_day_memos": [{"date": d, "memo": m} for d, m in sorted(day_memos.items(), reverse=True)],
        "recent_activities": [
            {"date": a["date"], "ex": a["ex"], "km": a["km"],
             "min": round(a["sec"] / 60) if a["sec"] else None, "level": a["level"], "memo": a["memo"]}
            for a in data["activities"] if since <= a["date"] < target_date],
        "done_on_target_date": [{"ex": r["ex"], "sets": r["sets"]} for r in done_today],
    }


SYSTEM = """당신은 일반인 헬스 이용자를 돕는 보수적인 근력운동 코치다. 의학적 진단이나 치료 조언은 하지 않는다.
사용자 메시지는 target_date 하루의 운동 계획을 짜기 위한 JSON 이다. stats 의 숫자는 이미 계산된 값이니 그대로 믿는다.
- exercises 목록에 있는 운동만 쓴다.
- 진행 규칙(더블 프로그레션): stats 의 hit_top_of_range 가 true 면(지난번 모든 세트가 rep 범위 상단 도달) 그 운동의 step 만큼만 올리고,
  아니면 무게는 그대로 두고 횟수를 채운다. 한 번에 step 보다 크게 올리지 않는다. 지난 기록이 없는 운동은 kg 를 null 로 둔다.
- profile.injuries, 운동별 memos, recent_day_memos, request.note 에 통증·부상·컨디션 저하가 있으면 관련 부위는 증량하지 않고,
  부담 큰 운동은 빼거나 가벼운 대안으로 바꾸고, 그 이유를 warnings 에 적는다.
- request.focus 부위를 우선하되(null 이면 part_last_trained 를 보고 가장 오래 쉰 부위 위주), 최근 48시간 안에 한 부위는 피한다.
- done_on_target_date 에 있는 운동은 그날 이미 한 것이니 다시 넣지 않는다.
- request.time_budget_min 안에 들도록 운동 수·세트 수를 정한다 (세트당 대략 휴식시간 + 40초, 워밍업 세트 포함).
- kg 는 그 운동에서 실제로 쓰던 단위(stats 의 지난 무게, step 의 배수)를 따른다. warm 은 지난 워밍업을 참고하되 필요 없으면 null.
- reps 는 목표 횟수(정수), repsText 는 범위로 보여줄 때만 '12~15' 처럼, 아니면 빈 문자열. rest 는 초.
- plan.note·items[].note·rationale[].why 는 한국어로 짧게, 근거가 된 지난 기록(날짜·무게·횟수)을 언급한다."""


def plan_schema(ex_names):
    """구조화 출력용 JSON 스키마. 숫자 범위(sets 1~6 등)는 API 가 안 받으니 validate() 에서 파이썬이 강제."""
    obj = lambda props: {"type": "object", "properties": props, "required": list(props), "additionalProperties": False}
    num_or_null = {"anyOf": [{"type": "number"}, {"type": "null"}]}
    item = obj({
        "ex": {"type": "string", "enum": ex_names},
        "warm": {"anyOf": [obj({"kg": {"type": "number"}, "reps": {"type": "integer"}}), {"type": "null"}]},
        "kg": num_or_null,
        "reps": {"type": "integer"},
        "repsText": {"type": "string"},
        "sets": {"type": "integer"},
        "rest": {"type": "integer"},
        "note": {"type": "string"},
    })
    return obj({
        "plan": obj({"title": {"type": "string"}, "note": {"type": "string"}, "items": {"type": "array", "items": item}}),
        "rationale": {"type": "array", "items": obj({"ex": {"type": "string"}, "why": {"type": "string"}})},
        "warnings": {"type": "array", "items": {"type": "string"}},
    })


def call_llm(ctx, model=MODEL):
    """컨텍스트 → (LLM 출력 dict, 메타{model, stop_reason, usage, ms})"""
    import anthropic                        # --context 만 쓸 때는 SDK 없어도 되게
    client = anthropic.Anthropic()          # ANTHROPIC_API_KEY 환경변수
    t0 = time.time()
    r = client.messages.create(
        model=model,
        max_tokens=MAX_TOKENS,
        system=SYSTEM,
        messages=[{"role": "user", "content": json.dumps(ctx, ensure_ascii=False, separators=(",", ":"))}],
        output_config={"format": {"type": "json_schema", "schema": plan_schema([e["name"] for e in ctx["exercises"]])}},
    )
    meta = {"model": r.model, "stop_reason": r.stop_reason, "ms": int((time.time() - t0) * 1000),
            "usage": {"input": r.usage.input_tokens, "output": r.usage.output_tokens}}
    if r.stop_reason != "end_turn":         # refusal·max_tokens 면 스키마대로가 아닐 수 있음
        raise RuntimeError(f"LLM 응답이 끝까지 오지 않음: stop_reason={r.stop_reason}")
    text = next(b.text for b in r.content if b.type == "text")
    return json.loads(text), meta


def to_plan(out, target_date):
    """LLM 출력 → plan.json 의 계획 하나 (앱 형식). warm 이 null 이면 빼서 '지난번 워밍업대로', repsText 가 비면 뺌"""
    items = []
    for it in out["plan"]["items"]:
        it = dict(it)
        if it.get("warm") is None:
            it.pop("warm", None)
        if not it.get("repsText"):
            it.pop("repsText", None)
        items.append(it)
    return {"id": f"ai-{target_date}", "date": target_date, "title": out["plan"]["title"],
            "note": out["plan"]["note"], "items": items}


def main():
    ap = argparse.ArgumentParser(description="LLM 운동 코치 — 다음 운동 계획 → plan.json")
    ap.add_argument("--date", default=datetime.date.today().isoformat(), help="계획 날짜 YYYY-MM-DD (기본: 오늘)")
    ap.add_argument("--focus", help="우선 부위 (예: 하체). 비우면 코치가 고름")
    ap.add_argument("--note", default="", help="오늘 컨디션·요청 (예: '허리 약간 뻐근')")
    ap.add_argument("--time", type=int, help="시간 예산(분). 기본: coach_profile.json")
    ap.add_argument("--context", action="store_true", help="LLM 없이 컨텍스트만 출력")
    ap.add_argument("--dry-run", action="store_true", help="계획을 출력만 하고 plan.json 은 안 씀")
    ap.add_argument("--model", default=MODEL, help=f"기본 {MODEL}")
    a = ap.parse_args()
    datetime.date.fromisoformat(a.date)       # 형식 검사

    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    ctx = build_context(sync_excel.read_excel(), a.date, a.focus, a.note, a.time, load_profile())
    if a.context:
        print(json.dumps(ctx, ensure_ascii=False, indent=2))
        return

    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        sys.exit("API 키가 없어요. setx ANTHROPIC_API_KEY \"<키>\" 후 새 터미널에서 실행하세요.")
    import anthropic
    try:
        out, meta = call_llm(ctx, a.model)
    except anthropic.AuthenticationError:
        sys.exit("API 키가 틀렸어요. 환경변수 ANTHROPIC_API_KEY 를 확인하고 새 터미널에서 다시 실행하세요.")
    except anthropic.APIConnectionError:
        sys.exit("네트워크 오류 — 인터넷 연결 확인")
    except anthropic.APIStatusError as e:
        sys.exit(f"API 오류 {e.status_code}: {e.message}")
    print(f"※ {meta['model']} · 입력 {meta['usage']['input']} / 출력 {meta['usage']['output']} 토큰 · {meta['ms'] / 1000:.1f}초",
          file=sys.stderr)
    print(json.dumps({"plan": to_plan(out, a.date), "rationale": out["rationale"], "warnings": out["warnings"]},
                     ensure_ascii=False, indent=2))
    if not a.dry_run:
        print("※ plan.json 저장은 아직 없음 (§8-5). 지금은 --dry-run 과 같음", file=sys.stderr)


if __name__ == "__main__":
    main()
