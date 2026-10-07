# -*- coding: utf-8 -*-
"""validate() 가드레일 단위 테스트 — 가짜 엑셀 데이터 + 가짜 LLM 출력으로 규칙별 확인.
   실행: python -m unittest discover -s tests -v"""
import os, sys, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import coach

TARGET = "2026-10-08"


def rec(date, ex, kg, reps, n=3, memo="", warm=None):
    return {"date": date, "ex": ex, "memo": memo, "warm": warm, "sets": [{"kg": kg, "reps": reps}] * n,
            "start": None, "end": None}


EXERCISES = [
    {"n": "레그컬", "p": "하체", "t": "머신", "step": 7.5, "rep": "10-15"},
    {"n": "레그익스텐션", "p": "하체", "t": "머신", "step": 3.75, "rep": "12-15"},
    {"n": "랫풀다운", "p": "등", "t": "머신", "step": 3.75, "rep": "10-15"},
    {"n": "바벨 컬", "p": "팔", "t": "프리웨이트", "step": 5, "rep": "8-12"},
    {"n": "플랭크", "p": "코어", "t": "맨몸", "step": 0, "rep": "30-60초"},
]
RECORDS = [
    rec("2026-10-01", "레그컬", 40, 15),                                   # 상단 15 도달 → +7.5 까지
    rec("2026-10-01", "레그익스텐션", 37.5, 10),                            # 상단 12 미달 → 유지
    rec("2026-10-01", "랫풀다운", 40, 15, memo="2세트에서 왼쪽 어깨 통증"),  # 도달했지만 통증 메모
    rec("2026-10-03", "바벨 컬", 15, 12, memo="팔꿈치 통증 없음"),           # 도달, 부정문 메모
    rec("2026-10-03", "플랭크", 0, 60, n=2),
]


def ctx(note="", records=RECORDS, injuries=(), budget=60, hold=()):
    data = {"exercises": EXERCISES, "records": list(records), "activities": []}
    profile = {**coach.DEFAULT_PROFILE, "injuries": list(injuries), "hold": list(hold)}
    return coach.build_context(data, TARGET, note=note, time_budget_min=budget, profile=profile)


def item(ex, kg, reps=12, sets=3, rest=60, warm=None, note=""):
    return {"ex": ex, "warm": warm, "kg": kg, "reps": reps, "repsText": "", "sets": sets, "rest": rest, "note": note}


def run(items, **kw):
    out = {"plan": {"title": "t", "note": "n", "items": items}, "rationale": [], "warnings": []}
    res, fixes = coach.validate(out, ctx(**kw))
    return res, [r for r, _ in fixes], {it["ex"]: it for it in res["plan"]["items"]}


class TestContext(unittest.TestCase):
    def test_next_kg_max(self):
        s = {x["ex"]: x for x in ctx()["stats"]}
        self.assertEqual(s["레그컬"]["next_kg_max"], 47.5)          # 도달 → +step
        self.assertEqual(s["레그익스텐션"]["next_kg_max"], 37.5)    # 미달 → 유지
        self.assertTrue(s["바벨 컬"]["hit_top_of_range"])


class TestValidate(unittest.TestCase):
    def test_ok_plan_unchanged(self):
        res, rules, it = run([item("레그컬", 47.5), item("레그익스텐션", 37.5)])
        self.assertEqual(rules, [])
        self.assertEqual(it["레그컬"]["kg"], 47.5)
        self.assertFalse(res["plan"]["note"].startswith("⚠"))

    def test_g1_unknown_exercise_removed(self):
        res, rules, it = run([item("스미스 스쿼트", 40), item("레그컬", 40)])
        self.assertIn("G1", rules)
        self.assertNotIn("스미스 스쿼트", it)
        self.assertIn("레그컬", it)

    def test_g2_cap_at_last_plus_step(self):
        res, rules, it = run([item("레그컬", 60)])
        self.assertIn("G2", rules)
        self.assertEqual(it["레그컬"]["kg"], 47.5)
        self.assertTrue(it["레그컬"]["note"].startswith("[자동수정]"))
        self.assertTrue(any(w.startswith("[자동수정] G2") for w in res["warnings"]))

    def test_g7_hold_weight_when_top_not_reached(self):
        res, rules, it = run([item("레그익스텐션", 41.25)])         # 10/4 실제 Haiku 출력과 같은 위반
        self.assertEqual(rules, ["G7"])                             # +step 이라 G2 는 통과, G7 이 잡음
        self.assertEqual(it["레그익스텐션"]["kg"], 37.5)

    def test_g3_request_note_freezes_body_part(self):
        res, rules, it = run([item("레그컬", 47.5), item("바벨 컬", 20)], note="허리 약간 뻐근")
        self.assertIn("G3", rules)
        self.assertEqual(it["레그컬"]["kg"], 40)                    # 허리 → 하체 증량 0
        self.assertEqual(it["바벨 컬"]["kg"], 20)                   # 팔은 상관없음

    def test_g3_unspecified_pain_freezes_all(self):
        res, rules, it = run([item("바벨 컬", 20)], note="컨디션 안 좋고 여기저기 아픔")
        self.assertEqual(it["바벨 컬"]["kg"], 15)

    def test_g3_own_memo_pain(self):
        res, rules, it = run([item("랫풀다운", 43.75)])
        self.assertIn("G3", rules)
        self.assertEqual(it["랫풀다운"]["kg"], 40)

    def test_g3_negated_memo_allows_increase(self):
        res, rules, it = run([item("바벨 컬", 20)])                 # '팔꿈치 통증 없음'
        self.assertEqual(rules, [])
        self.assertEqual(it["바벨 컬"]["kg"], 20)

    def test_g3_profile_injuries_do_not_block(self):
        res, rules, it = run([item("레그컬", 47.5)], injuries=["허리 디스크 주의"])
        self.assertEqual(rules, [])

    def test_g3_recent_day_memo(self):
        recs = RECORDS + [rec("2026-10-05", "랫풀다운", 40, 12, memo="[오늘] 왼 무릎 시림 · 가볍게")]
        res, rules, it = run([item("레그컬", 47.5)], records=recs)
        self.assertEqual(it["레그컬"]["kg"], 40)

    def test_g3_old_day_memo_ignored(self):
        recs = RECORDS + [rec("2026-09-20", "랫풀다운", 30, 12, memo="[오늘] 무릎 통증")]
        res, rules, it = run([item("레그컬", 47.5)], records=recs)
        self.assertEqual(it["레그컬"]["kg"], 47.5)

    def test_g4_clamp_sets_reps(self):
        res, rules, it = run([item("레그컬", 40, reps=45, sets=9), item("레그익스텐션", 37.5, reps=0, sets=0)])
        self.assertEqual((it["레그컬"]["sets"], it["레그컬"]["reps"]), (6, 30))
        self.assertEqual((it["레그익스텐션"]["sets"], it["레그익스텐션"]["reps"]), (1, 1))

    def test_g4_plank_seconds_not_clipped(self):
        res, rules, it = run([item("플랭크", 0, reps=60, sets=2)])
        self.assertEqual(it["플랭크"]["reps"], 60)
        self.assertNotIn("G4", rules)

    def test_g5_time_budget_cuts_from_last(self):
        # 3개 × (5세트 + 지난 워밍업 없음) × (60+40초) = 25분, 예산 10분 × 1.2 = 12분
        items = [item("레그컬", 40, sets=5), item("레그익스텐션", 37.5, sets=5), item("바벨 컬", 15, sets=5)]
        res, rules, it = run(items, budget=10)
        self.assertIn("G5", rules)
        stats = {s["ex"]: s for s in ctx()["stats"]}
        self.assertLessEqual(coach.est_minutes(res["plan"]["items"], stats), 12)
        self.assertEqual(it["바벨 컬"]["sets"], 1)                  # 뒤 항목부터 깎임
        self.assertGreater(it["레그컬"]["sets"], it["바벨 컬"]["sets"])

    def test_g6_done_today_warning_only(self):
        recs = RECORDS + [rec(TARGET, "바벨 컬", 15, 10)]
        res, rules, it = run([item("바벨 컬", 15)], records=recs)
        self.assertEqual(rules, ["G6"])
        self.assertIn("바벨 컬", it)                                 # 지우지 않음
        self.assertTrue(any(w.startswith("[주의]") for w in res["warnings"]))

    def test_null_kg_skips_weight_rules(self):
        res, rules, it = run([item("레그컬", None)], note="허리 뻐근")
        self.assertEqual(rules, [])


HOLD = [{"exercises": ["랫풀다운"], "reason": "왼 어깨 — 진단 전까지 제외"}]


class TestHold(unittest.TestCase):
    def test_held_exercise_not_offered(self):
        c = ctx(hold=HOLD)
        names = [e["name"] for e in c["exercises"]]
        self.assertNotIn("랫풀다운", names)                       # 스키마 enum 에도 안 들어감
        self.assertNotIn("랫풀다운", coach.plan_schema(names)["properties"]["plan"]["properties"]["items"]["items"]["properties"]["ex"]["enum"])
        self.assertEqual(c["holds"], [{"reason": "왼 어깨 — 진단 전까지 제외", "exercises": ["랫풀다운"]}])

    def test_hold_by_part(self):
        c = ctx(hold=[{"parts": ["팔"], "reason": "팔꿈치"}])
        self.assertNotIn("바벨 컬", [e["name"] for e in c["exercises"]])
        self.assertIn("레그컬", [e["name"] for e in c["exercises"]])

    def test_g8_removes_held(self):
        res, rules, it = run([item("랫풀다운", 40), item("레그컬", 40)], hold=HOLD)
        self.assertEqual(rules, ["G8"])                            # G1(목록에 없음)이 아니라 G8 로
        self.assertNotIn("랫풀다운", it)
        self.assertTrue(any("진단 전까지" in w for w in res["warnings"]))

    def test_no_hold_default(self):
        self.assertEqual(ctx()["holds"], [])


class TestPain(unittest.TestCase):
    def test_has_pain(self):
        self.assertTrue(coach.has_pain("왼쪽 무릎 통증 → 50으로 낮춤"))
        self.assertFalse(coach.has_pain("무릎·허리 통증 없었음"))
        self.assertFalse(coach.has_pain("위에서 1초 멈춤 처음 적용"))

    def test_illness_freezes_all(self):
        self.assertEqual(coach.pain_parts("감기 기운"), coach.ALL_PARTS)
        self.assertEqual(coach.pain_parts("몸살 기운에 허리도 뻐근"), coach.ALL_PARTS)   # 부위 단어가 있어도 전 부위
        self.assertEqual(coach.pain_parts("감기 없음"), set())

    def test_illness_note_blocks_increase(self):
        res, rules, it = run([item("레그컬", 47.5), item("바벨 컬", 20)], note="아침에 감기 기운")
        self.assertEqual((it["레그컬"]["kg"], it["바벨 컬"]["kg"]), (40, 15))

    def test_pain_parts(self):
        self.assertEqual(coach.pain_parts("허리 약간 뻐근"), {"하체", "등", "코어"})
        self.assertEqual(coach.pain_parts("팔꿈치 통증 없음"), set())


if __name__ == "__main__":
    unittest.main()
