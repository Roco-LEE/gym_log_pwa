# -*- coding: utf-8 -*-
"""LLM 운동 코치 (1단계) — 엑셀 기록을 읽어 다음 운동 계획을 짜고 plan.json 에 넣는다.
   설계: docs/LLM코치_1단계_설계.md

   python coach.py --date 2026-10-08 --focus 하체 --note "허리 약간 뻐근" --context
       → LLM 없이 컨텍스트(LLM 에 보낼 JSON)만 출력. 통계가 맞는지 눈으로 확인용
   python coach.py --date 2026-10-08 --focus 하체 --note "허리 약간 뻐근" --dry-run
       → LLM 으로 계획을 받아 출력만 (plan.json 은 안 씀)
   python coach.py --date 2026-10-08 --focus 하체 --note "허리 약간 뻐근"
       → 계획을 plan.json 에 병합 저장 (id ai-날짜, 같은 id 는 교체·손으로 쓴 계획은 유지, 쓰기 전 백업/plan_*.json)

- 통계(지난 세트·증량 신호·e1RM·부위별 마지막 날짜)는 파이썬이 계산해서 준다. LLM 에게 계산을 시키지 않는다.
- 개인 설정(목표·부상·주당 횟수·시간)은 coach_profile.json (git 제외). 없으면 기본값.
- LLM 출력은 구조화 출력(output_config.format = JSON 스키마)으로 받는다. 운동 이름은 스키마 enum.
  (설계의 '도구 강제 호출'은 Sonnet 5.5·Opus 5.5 에서 400 이라, 모델을 바꿔도 그대로 되는 쪽으로)
- API 키는 OS 환경변수 ANTHROPIC_API_KEY 에만. 이 폴더에 두지 않는다 (serve.py 가 폴더를 서빙함).
- 호출마다 coach_log/ 에 입력 컨텍스트·LLM 원문·가드레일 수정·최종 계획·토큰·비용·지연을 남긴다 (git 제외, 서빙 차단).
"""
import argparse, datetime, glob, hashlib, json, os, re, shutil, sys, time
import sync_excel

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.path.join(HERE, "coach_profile.json")
PLAN = os.path.join(HERE, "plan.json")
PLAN_BACKUPS = 10            # 백업 폴더에 남길 plan_*.json 수
# 2026-10-06 Haiku 4.5 와 비교: Haiku 는 허리 뻐근한데 RDL 을 넣는 등 판단·사실 오류 → Sonnet 5.5 (1회 약 5센트)
MODEL = "claude-sonnet-5-5"
MAX_TOKENS = 16000
FALLBACK_BETA = "server-side-fallback-2026-07-01"
FALLBACK_MODELS = {"claude-sonnet-5-5", "claude-opus-5-5", "claude-opus-5", "claude-fable-5-1"}
WINDOW_DAYS = 42          # 최근 범위: 6주 또는 최근 12세션 중 짧은 쪽
WINDOW_SESSIONS = 12
MEMOS_PER_EX = 3          # 운동별로 보여줄 최근 메모 수
DAY_TAG = "[오늘] "
DEFAULT_PROFILE = {"goal": "근비대·자세 우선", "injuries": [], "hold": [], "days_per_week": 3, "time_budget_min": 60}
# injuries = 늘 안고 가는 지병·이력 (LLM 맥락용, 증량 금지 근거로는 안 씀)
# hold     = 진단·회복 전까지 아예 빼는 운동. [{"exercises": [이름…], "parts": [부위…], "reason": "…"}]
#            → 컨텍스트의 운동 목록(스키마 enum)에서 빠지고, 그래도 들어오면 G8 이 지움


def held_exercises(exercise_rows, profile):
    """보류 중인 운동 {이름: 이유}. exercise_rows 는 read_excel() 의 exercises ({n, p, …})"""
    held = {}
    for h in profile.get("hold") or []:
        names, parts = set(h.get("exercises") or []), set(h.get("parts") or [])
        for e in exercise_rows:
            if e["n"] in names or (e["p"] and e["p"] in parts):
                held.setdefault(e["n"], h.get("reason") or "보류")
    return held


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
    # 보류 운동도 계획 후보에서 뺌 (스키마 enum 에도 안 들어감)
    held = held_exercises(data["exercises"], profile)
    exercises = [{"name": e["n"], "part": e["p"], "type": e["t"], "step": e["step"], "rep_range": e["rep"]}
                 for e in data["exercises"] if e["p"] and e["n"] not in held]
    holds = {}
    for name, reason in held.items():
        holds.setdefault(reason, []).append(name)

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
        top = max((s["kg"] for s in last["sets"]), default=None)
        hit = bool(rng and reps and all(x >= rng[1] for x in reps))
        stats.append({
            "ex": ex,
            "part": info.get("p", ""),
            "last_date": last["date"],
            "days_since": (tgt - datetime.date.fromisoformat(last["date"])).days,
            "last_warm": last["warm"],
            "last_sets": last["sets"],
            "last_top_kg": top,
            "hit_top_of_range": hit,
            # 더블 프로그레션으로 이번에 쓸 수 있는 최대 무게 (LLM 이 계산하지 않게)
            "next_kg_max": None if top is None else round(top + (info.get("step") or 0), 2) if hit else top,
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
        "holds": [{"reason": r, "exercises": names} for r, names in holds.items()],
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
- holds 의 운동은 진단·회복 전까지 보류 중이라 목록에서 빠져 있다. 쓰지 말고, 그 부위에 부담이 가는 비슷한 동작도 피하며,
  보류 중이라는 사실을 warnings 에 한 줄 적는다.
- request.note·recent_day_memos 에 감기·몸살 등 컨디션 저하가 있으면 전 부위 증량하지 않고 세트·운동 수를 줄인다.
- 진행 규칙(더블 프로그레션): kg 는 stats 의 next_kg_max 를 절대 넘지 않는다. next_kg_max 는 지난번 모든 세트가 rep 범위 상단에
  도달했으면(hit_top_of_range=true) 지난 무게 + step, 아니면 지난 무게 그대로다. 무게를 유지하는 운동은 횟수를 채우는 게 목표다.
  지난 기록이 없는 운동은 kg 를 null 로 둔다.
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
    """컨텍스트 → (LLM 이 낸 원문 텍스트, 메타{model, stop_reason, usage, ms}). 파싱·stop_reason 판단은 호출한 쪽에서"""
    import anthropic                        # --context 만 쓸 때는 SDK 없어도 되게
    client = anthropic.Anthropic()          # ANTHROPIC_API_KEY 환경변수
    # 안전 분류기가 거절하면 서버가 권장 모델로 다시 돌림 (Sonnet 5.5 등 최신 모델만 지원, Haiku 는 안 붙임)
    fb = {"betas": [FALLBACK_BETA], "fallbacks": "default"} if model in FALLBACK_MODELS else {}
    t0 = time.time()
    r = client.beta.messages.create(
        model=model,
        max_tokens=MAX_TOKENS,
        system=SYSTEM,
        messages=[{"role": "user", "content": json.dumps(ctx, ensure_ascii=False, separators=(",", ":"))}],
        output_config={"format": {"type": "json_schema", "schema": plan_schema([e["name"] for e in ctx["exercises"]])}},
        **fb,
    )
    meta = {"model": r.model, "stop_reason": r.stop_reason, "ms": int((time.time() - t0) * 1000),
            "usage": {"input": r.usage.input_tokens, "output": r.usage.output_tokens}}
    text = "".join(b.text for b in r.content if b.type == "text")
    return text, meta


# ===================== 호출 로그 (coach_log/, git 제외) =====================
# 나중에 Eval 정답 세트 재료: 같은 입력(context)으로 프롬프트·모델을 바꿔 다시 돌려 비교할 수 있게 전부 남김
LOG_DIR = os.path.join(HERE, "coach_log")
PRICES = {"claude-sonnet-5-5": (2, 10), "claude-haiku-4-5": (1, 5), "claude-opus-5-5": (4, 20)}   # $ / 1M 토큰 (입력, 출력)


def cost_usd(model, usage):
    """모델 ID 앞부분으로 단가를 찾아 대략 비용($). 표에 없으면 None"""
    price = next((v for k, v in PRICES.items() if (model or "").startswith(k)), None)
    if not price or not usage:
        return None
    return round((usage["input"] * price[0] + usage["output"] * price[1]) / 1e6, 5)


def prompt_version():
    """시스템 프롬프트·스키마 틀이 바뀌면 달라지는 짧은 해시 — 로그끼리 비교할 때 같은 프롬프트였는지"""
    return hashlib.sha256((SYSTEM + json.dumps(plan_schema(["x"]), sort_keys=True)).encode("utf-8")).hexdigest()[:10]


def write_log(entry, log_dir=None):
    """호출 한 번 = coach_log/YYYYMMDD-HHMMSS_<계획날짜>.json 하나. 반환: 경로"""
    log_dir = log_dir or LOG_DIR
    os.makedirs(log_dir, exist_ok=True)
    path = os.path.join(log_dir, f"{datetime.datetime.now():%Y%m%d-%H%M%S}_{entry['target_date']}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(entry, f, ensure_ascii=False, indent=2)
    return path


# ===================== 가드레일 (LLM 뒤에서 파이썬이 강제) =====================
MAX_SETS, MAX_REPS = 6, 30          # [기록] 시트가 6세트까지
DEFAULT_REST = 90
SEC_PER_SET = 40                    # 세트 수행 시간 대략
PAIN = re.compile(r"허리|통증|아픔|아프|아팠|뻐근|부상|시림|시린|감기|몸살|발열|오한|독감")
ILLNESS = re.compile(r"감기|몸살|발열|오한|독감")   # 몸 전체 컨디션 → 부위 단어가 있어도 전 부위
BODY_PARTS = {"허리": {"하체", "등", "코어"}, "무릎": {"하체"}, "어깨": {"어깨", "가슴", "등"},
              "팔꿈치": {"팔", "가슴"}, "손목": {"팔", "가슴"}}
ALL_PARTS = {"하체", "등", "코어", "어깨", "가슴", "팔"}
PAIN_MEMO_DAYS = 7                  # 그날 메모는 최근 7일 것만 G3 근거로


def has_pain(text):
    """통증·부상 키워드가 있나. '통증 없음'·'허리 통증 없었음' 처럼 바로 뒤에 '없'이 오면 부정으로 보고 넘김"""
    text = text or ""
    return any("없" not in text[m.end():m.end() + 8] for m in PAIN.finditer(text))


def pain_parts(text):
    """통증 문장이 가리키는 부위들. 몸 부위 단어가 없으면 전 부위(보수적)"""
    if not has_pain(text):
        return set()
    if any("없" not in text[m.end():m.end() + 8] for m in ILLNESS.finditer(text)):
        return set(ALL_PARTS)
    parts = set().union(*(v for k, v in BODY_PARTS.items() if k in text))
    return parts or set(ALL_PARTS)


def est_minutes(items, stats_by_ex):
    """예상 시간(분): (본 세트 + 워밍업 1) × (휴식 + 40초). warm 이 없으면 앱이 지난번 워밍업을 쓰므로 그것도 셈"""
    sec = 0
    for it in items:
        warm = it.get("warm")
        if warm is None:
            warm = (stats_by_ex.get(it["ex"]) or {}).get("last_warm")
        sec += (it["sets"] + (1 if warm else 0)) * ((it.get("rest") or DEFAULT_REST) + SEC_PER_SET)
    return sec / 60


def validate(out, ctx):
    """LLM 출력(out)을 규칙대로 고친 사본과 고친 내역 [(규칙, 내용)] 을 돌려줌.
    고친 건 항목 note 앞에 [자동수정], warnings 에 '[자동수정] …' 으로 붙여 앱 카드에서 보이게 한다.
    G1 운동목록에 없는 운동 → 삭제
    G2 kg ≤ 지난번 최고 무게 + step
    G3 통증 메모(오늘 요청·최근 7일 그날 메모·그 운동의 마지막 메모)가 가리키는 부위는 증량 0
       (profile.injuries 는 늘 있는 지병이라 근거로 안 씀 — 쓰면 영영 증량이 안 됨. LLM 에게 맥락으로만 줌)
    G4 sets 1~6, reps 1~30 (플랭크처럼 범위가 30 넘는 운동은 그 상단까지)
    G5 예상 시간 ≤ 시간 예산 × 1.2 → 뒤 항목부터 세트 축소
    G6 그날 이미 한 운동과 겹침 → 경고만
    G7 지난번에 rep 범위 상단을 못 채웠으면 무게 유지 (kg ≤ next_kg_max) — G2 만으론 '미달인데 +step' 을 못 잡음
    G8 profile.hold 로 보류 중인 운동 → 삭제 (enum 에서 빠져 있어 보통은 안 들어옴, 안전망)"""
    out = json.loads(json.dumps(out))       # 깊은 복사
    fixes = []
    ex_info = {e["name"]: e for e in ctx["exercises"]}
    stats = {s["ex"]: s for s in ctx["stats"]}
    tgt = datetime.date.fromisoformat(ctx["target_date"])

    def fix(it, rule, msg):
        fixes.append((rule, f"{it['ex']}: {msg}" if it else msg))
        if it is not None:
            it["note"] = f"[자동수정] {msg}" + (" · " + it["note"] if it.get("note") else "")

    # G3 근거: 오늘 요청 + 최근 7일 그날 메모 → 부위 단위
    frozen = pain_parts(ctx["request"].get("note"))
    for m in ctx.get("recent_day_memos", []):
        if (tgt - datetime.date.fromisoformat(m["date"])).days <= PAIN_MEMO_DAYS:
            frozen |= pain_parts(m["memo"])

    items = []
    for it in out["plan"]["items"]:
        hold = next((h["reason"] for h in ctx.get("holds", []) if it["ex"] in h["exercises"]), None)
        if hold:                                                             # G8
            fix(None, "G8", f"보류 중인 '{it['ex']}' 삭제 ({hold})")
            continue
        info = ex_info.get(it["ex"])
        if not info:                                                         # G1
            fix(None, "G1", f"운동목록에 없는 '{it['ex']}' 삭제")
            continue
        st = stats.get(it["ex"])

        rng = rep_range(info.get("rep_range"))                               # G4
        rmax = max(MAX_REPS, rng[1] if rng else 0)
        if not 1 <= it["sets"] <= MAX_SETS:
            new = min(max(it["sets"], 1), MAX_SETS)
            fix(it, "G4", f"세트 {it['sets']}→{new}")
            it["sets"] = new
        if not 1 <= it["reps"] <= rmax:
            new = min(max(it["reps"], 1), rmax)
            fix(it, "G4", f"횟수 {it['reps']}→{new}")
            it["reps"] = new
        if it.get("kg") is not None and it["kg"] < 0:
            fix(it, "G4", f"무게 {it['kg']}→0")
            it["kg"] = 0

        if it.get("kg") is not None and st and st["last_top_kg"] is not None:
            last = st["last_top_kg"]
            cap2 = last + (info.get("step") or 0)                            # G2
            if it["kg"] > cap2:
                fix(it, "G2", f"{it['kg']}kg→{cap2:g}kg (한 번에 step {info.get('step'):g}kg 까지만)")
                it["kg"] = cap2
            if not st["hit_top_of_range"] and it["kg"] > last:               # G7
                fix(it, "G7", f"{it['kg']:g}kg→{last:g}kg (지난번 상단 {rng[1] if rng else '?'}회 미달 → 무게 유지)")
                it["kg"] = last
            own = st["memos"][0]["memo"] if st["memos"] and st["memos"][0]["date"] == st["last_date"] else ""
            if it["kg"] > last and (info["part"] in frozen or has_pain(own)):  # G3
                why = "그 운동 지난 메모" if has_pain(own) else "통증·컨디션 메모"
                fix(it, "G3", f"{it['kg']:g}kg→{last:g}kg ({why} 있어 증량 보류)")
                it["kg"] = last

        if any(d["ex"] == it["ex"] for d in ctx.get("done_on_target_date", [])):  # G6
            fixes.append(("G6", f"{it['ex']}: 그날 이미 한 운동과 겹침"))
        items.append(it)

    budget = ctx["request"]["time_budget_min"] * 1.2                         # G5
    before = est_minutes(items, stats)
    while est_minutes(items, stats) > budget:
        cut = next((it for it in reversed(items) if it["sets"] > 1), None)
        if not cut:
            break
        cut["sets"] -= 1
        fixes.append(("G5", f"{cut['ex']}: 시간 초과로 세트 -1 → {cut['sets']}세트"))
    if est_minutes(items, stats) < before:
        fixes.append(("G5", f"예상 {before:.0f}분 → {est_minutes(items, stats):.0f}분 (예산 {ctx['request']['time_budget_min']}분)"))

    out["plan"]["items"] = items
    out["warnings"] = out.get("warnings", []) + [f"[자동수정] {r} {m}" if r != "G6" else f"[주의] {m}" for r, m in fixes]
    if fixes:
        out["plan"]["note"] = f"⚠ 자동수정 {sum(r != 'G6' for r, _ in fixes)}건 · " + out["plan"]["note"]
    return out, fixes


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


def merge_plans(doc, plan):
    """plan.json 내용(doc)에 계획 하나를 넣은 새 doc 과 '교체'/'추가'.
    같은 id(ai-날짜)는 그 자리에서 교체, 손으로 쓴 계획·지난 날짜 계획은 그대로 둠"""
    plans = list((doc or {}).get("plans") or [])
    i = next((k for k, p in enumerate(plans) if p.get("id") == plan["id"]), None)
    if i is None:
        plans.append(plan)
    else:
        plans[i] = plan
    return {**(doc or {}), "plans": plans}, "추가" if i is None else "교체"


def save_plan(plan, path=None, backup_dir=None):
    """plan.json 에 병합 저장. 쓰기 전 기존 파일을 백업 폴더에 plan_YYYYMMDD-HHMMSS.json 으로 (최근 10개).
    반환: ('교체'|'추가', 백업 경로 또는 None)"""
    path = path or PLAN
    backup_dir = backup_dir or sync_excel.BACKUP_DIR
    doc, bak = None, None
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            doc = json.load(f)                  # 깨진 파일이면 여기서 멈춤 (덮어쓰지 않음)
        os.makedirs(backup_dir, exist_ok=True)
        bak = os.path.join(backup_dir, f"plan_{datetime.datetime.now():%Y%m%d-%H%M%S}.json")
        shutil.copy2(path, bak)
        for old in sorted(glob.glob(os.path.join(backup_dir, "plan_*.json")))[:-PLAN_BACKUPS]:
            os.remove(old)
    doc, action = merge_plans(doc, plan)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, path)                       # 중간에 끊겨도 반쪽짜리 plan.json 이 안 남게
    return action, bak


def list_plans(path=None):
    """plan.json 의 계획 목록 (없으면 [])"""
    path = path or PLAN
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return (json.load(f) or {}).get("plans") or []


def delete_plan(plan_id, path=None, backup_dir=None):
    """plan.json 에서 AI 계획(id ai-…) 하나 삭제. 쓰기 전 백업. 손으로 쓴 계획은 지우지 않음.
    반환: (지운 계획 또는 None, 백업 경로 또는 None). 폰은 다음에 앱을 열 때 plan.json 을 받아 카드가 사라짐"""
    if not plan_id.startswith("ai-"):
        raise CoachError("AI 계획(ai-…)만 지울 수 있어요")
    path = path or PLAN
    if not os.path.exists(path):
        return None, None
    with open(path, encoding="utf-8") as f:
        doc = json.load(f)
    plans = doc.get("plans") or []
    gone = next((p for p in plans if p.get("id") == plan_id), None)
    if gone is None:
        return None, None
    backup_dir = backup_dir or sync_excel.BACKUP_DIR
    os.makedirs(backup_dir, exist_ok=True)
    bak = os.path.join(backup_dir, f"plan_{datetime.datetime.now():%Y%m%d-%H%M%S}.json")
    shutil.copy2(path, bak)
    doc["plans"] = [p for p in plans if p.get("id") != plan_id]
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, path)
    return gone, bak


def save_holds(holds, path=None):
    """coach_profile.json 의 hold 만 바꿔 저장 (다른 항목은 그대로)"""
    path = path or PROFILE
    prof = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            prof = json.load(f)
    prof["hold"] = holds
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prof, f, ensure_ascii=False, indent=2)
        f.write("\n")


class CoachError(Exception):
    """사용자에게 그대로 보여줄 수 있는 실패 (키 없음·네트워크·응답 끊김 등)"""


def make_plan(date, focus=None, note="", time=None, dry_run=False, model=MODEL):
    """계획 하나 만들기 — CLI(main)와 PC 창(panel.py)이 같이 쓰는 입구.
    엑셀 읽기 → 컨텍스트 → LLM → 가드레일 → (저장). 성공·실패 모두 coach_log 에 남김.
    반환: {plan, fixes:[(규칙, 내용)], meta, saved:'교체'|'추가'|None, backup, log}. 실패는 CoachError"""
    datetime.date.fromisoformat(date)          # 형식 검사
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        raise CoachError("API 키가 없어요. 환경변수 ANTHROPIC_API_KEY 를 설정한 뒤 다시 실행하세요.")
    try:
        data = sync_excel.read_excel()
    except FileNotFoundError:
        raise CoachError(f"엑셀을 못 찾았어요: {sync_excel.XLSX}")
    ctx = build_context(data, date, focus, note, time, load_profile())
    log = {"ts": datetime.datetime.now().isoformat(timespec="seconds"), "target_date": date,
           "args": {"focus": focus, "note": note, "time": time, "dry_run": dry_run},
           "model_requested": model, "prompt_version": prompt_version(), "context": ctx}
    try:
        result = _run(ctx, date, dry_run, model, log)
    except Exception as e:                     # 예상 못 한 오류도 입력·원문은 남김
        log["error"] = str(e) if isinstance(e, CoachError) else repr(e)
        raise
    finally:
        log_path = write_log(log)
    result["log"] = log_path
    return result


def _run(ctx, date, dry_run, model, log):
    """LLM 호출 → 검증 → (저장). 진행하며 log 를 채움"""
    import anthropic
    try:
        text, meta = call_llm(ctx, model)
    except anthropic.AuthenticationError:
        raise CoachError("API 키가 틀렸어요. 환경변수 ANTHROPIC_API_KEY 를 확인하세요.")
    except anthropic.APIConnectionError:
        raise CoachError("네트워크 오류 — 인터넷 연결 확인")
    except anthropic.APIStatusError as e:
        raise CoachError(f"API 오류 {e.status_code}: {e.message}")
    meta["cost_usd"] = cost_usd(meta["model"], meta["usage"])
    log.update(meta=meta, raw_output=text)
    if meta["stop_reason"] != "end_turn":      # refusal·max_tokens 면 스키마대로가 아닐 수 있음
        raise CoachError(f"LLM 응답이 끝까지 오지 않음: stop_reason={meta['stop_reason']}")
    out = json.loads(text)
    log["llm_output"] = out

    out, fixes = validate(out, ctx)
    log["fixes"] = [{"rule": r, "msg": m} for r, m in fixes]
    plan = to_plan(out, date)
    # 앱은 모르는 필드는 무시함 → 근거·경고도 계획에 같이 남겨 둠 (카드엔 note 만 보임)
    plan["rationale"], plan["warnings"] = out["rationale"], out["warnings"]
    plan["model"] = meta["model"]
    log["final_plan"] = plan
    saved, bak = (None, None) if dry_run else save_plan(plan)
    log["saved"] = saved
    return {"plan": plan, "fixes": fixes, "meta": meta, "saved": saved, "backup": bak}


def main():
    ap = argparse.ArgumentParser(description="LLM 운동 코치 — 다음 운동 계획 → plan.json")
    ap.add_argument("--date", default=datetime.date.today().isoformat(), help="계획 날짜 YYYY-MM-DD (기본: 오늘)")
    ap.add_argument("--focus", help="우선 부위 (예: 하체). 비우면 코치가 고름")
    ap.add_argument("--note", default="", help="오늘 컨디션·요청 (예: '허리 약간 뻐근')")
    ap.add_argument("--time", type=int, help="시간 예산(분). 기본: coach_profile.json")
    ap.add_argument("--context", action="store_true", help="LLM 없이 컨텍스트만 출력")
    ap.add_argument("--dry-run", action="store_true", help="계획을 출력만 하고 plan.json 은 안 씀")
    ap.add_argument("--model", default=MODEL, help=f"기본 {MODEL}")
    ap.add_argument("--delete", action="store_true", help="그 날짜의 AI 계획(ai-날짜)을 plan.json 에서 삭제")
    a = ap.parse_args()
    datetime.date.fromisoformat(a.date)       # 형식 검사

    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    if a.delete:
        gone, bak = delete_plan(f"ai-{a.date}")
        print(f"※ 삭제: {gone['title']} (백업: {bak})" if gone else f"※ ai-{a.date} 계획이 없어요", file=sys.stderr)
        return
    if a.context:
        ctx = build_context(sync_excel.read_excel(), a.date, a.focus, a.note, a.time, load_profile())
        print(json.dumps(ctx, ensure_ascii=False, indent=2))
        return

    try:
        r = make_plan(a.date, a.focus, a.note, a.time, a.dry_run, a.model)
    except CoachError as e:
        sys.exit(str(e))
    m = r["meta"]
    print(f"※ {m['model']} · 입력 {m['usage']['input']} / 출력 {m['usage']['output']} 토큰 · "
          f"{m['ms'] / 1000:.1f}초 · 약 ${m['cost_usd']}", file=sys.stderr)
    for rule, msg in r["fixes"]:
        print(f"※ {rule} {msg}", file=sys.stderr)
    if not r["fixes"]:
        print("※ 가드레일: 고칠 것 없음", file=sys.stderr)
    print(json.dumps(r["plan"], ensure_ascii=False, indent=2))
    if a.dry_run:
        print("※ --dry-run: plan.json 은 안 씀", file=sys.stderr)
    else:
        print(f"※ plan.json 에 {r['plan']['id']} {r['saved']}" + (f" (이전 파일 백업: {r['backup']})" if r["backup"] else ""),
              file=sys.stderr)
    print(f"※ 로그: {r['log']}", file=sys.stderr)


if __name__ == "__main__":
    main()
