# -*- coding: utf-8 -*-
"""PC 창(panel.py) 결과 표시 테스트 — 창은 띄우지 않고 format_plan 만.
   실행: python -m unittest discover -s tests -v"""
import os, sys, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import panel

META = {"model": "claude-sonnet-5-5", "ms": 25700, "cost_usd": 0.05, "usage": {"input": 1, "output": 1}}


def result(items, fixes=(), saved="추가", warnings=()):
    plan = {"id": "ai-2026-10-08", "date": "2026-10-08", "title": "제목", "note": "메모", "items": items,
            "warnings": list(warnings)}
    return {"plan": plan, "fixes": list(fixes), "meta": META, "saved": saved, "backup": None, "log": "x"}


def text(r):
    return "".join(t for t, _ in panel.format_plan(r))


class TestFormatPlan(unittest.TestCase):
    def test_item_line(self):
        t = text(result([{"ex": "레그컬", "warm": {"kg": 22.5, "reps": 12}, "kg": 37.5, "reps": 12, "sets": 3, "note": "유지"}]))
        self.assertIn("레그컬  W 22.5×12 · 37.5kg × 12회 · 3세트", t)
        self.assertIn("   유지", t)

    def test_seconds_and_null_kg(self):
        t = text(result([{"ex": "플랭크", "kg": None, "reps": 40, "repsText": "30~60초", "sets": 2}]))
        self.assertIn("플랭크  30~60초 · 2세트", t)          # '초회' 아님, kg 없음

    def test_fixes_and_warnings(self):
        r = result([], fixes=[("G7", "레그컬: 40kg→37.5kg"), ("G6", "바벨 컬: 겹침")],
                   warnings=["[자동수정] G7 레그컬: 40kg→37.5kg", "허리 통증 시 중단"])
        t = text(r)
        self.assertIn("• G7 레그컬: 40kg→37.5kg", t)
        self.assertNotIn("G6", t)                              # 경고만인 G6 는 자동수정 목록에서 뺌
        self.assertIn("• 허리 통증 시 중단", t)
        self.assertEqual(t.count("레그컬: 40kg→37.5kg"), 1)     # LLM 경고 목록에 같은 줄 중복 안 됨

    def test_saved_label(self):
        self.assertIn("같은 날짜 계획을 교체", text(result([], saved="교체")))
        self.assertIn("미리보기 (저장 안 함)", text(result([], saved=None)))


if __name__ == "__main__":
    unittest.main()
