# -*- coding: utf-8 -*-
"""coach_log 테스트 — 가짜 엑셀·가짜 LLM 으로 main() 을 끝까지 돌려 로그 파일 내용을 확인 (API 비용 없음).
   실행: python -m unittest discover -s tests -v"""
import glob, json, os, sys, tempfile, unittest
from unittest import mock
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import coach
from test_validate import EXERCISES, RECORDS

DATA = {"exercises": EXERCISES, "records": RECORDS, "activities": []}
META = {"model": "claude-sonnet-5-5", "stop_reason": "end_turn", "ms": 1234, "usage": {"input": 8000, "output": 3000}}


def llm_text(kg):
    return json.dumps({"plan": {"title": "t", "note": "n", "items": [
        {"ex": "레그익스텐션", "warm": None, "kg": kg, "reps": 12, "repsText": "", "sets": 3, "rest": 60, "note": ""}]},
        "rationale": [], "warnings": []}, ensure_ascii=False)


class TestHelpers(unittest.TestCase):
    def test_cost(self):
        self.assertEqual(coach.cost_usd("claude-sonnet-5-5", {"input": 8000, "output": 3000}), 0.046)
        self.assertEqual(coach.cost_usd("claude-haiku-4-5-20251001", {"input": 1_000_000, "output": 0}), 1.0)
        self.assertIsNone(coach.cost_usd("other-model", {"input": 1, "output": 1}))

    def test_prompt_version_stable(self):
        self.assertEqual(coach.prompt_version(), coach.prompt_version())
        self.assertEqual(len(coach.prompt_version()), 10)


class TestMainLog(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.null = [open(os.devnull, "w", encoding="utf-8") for _ in range(2)]   # main() 이 reconfigure 를 부름
        self.patches = [
            mock.patch.object(coach, "LOG_DIR", os.path.join(self.tmp.name, "coach_log")),
            mock.patch.object(coach, "PLAN", os.path.join(self.tmp.name, "plan.json")),
            mock.patch.object(coach.sync_excel, "read_excel", return_value=DATA),
            mock.patch.object(coach.sync_excel, "BACKUP_DIR", os.path.join(self.tmp.name, "백업")),
            mock.patch.dict(os.environ, {"ANTHROPIC_API_KEY": "test"}),
            mock.patch("sys.stdout", new=self.null[0]),
            mock.patch("sys.stderr", new=self.null[1]),
        ]
        for p in self.patches:
            p.start()

    def tearDown(self):
        for p in reversed(self.patches):
            p.stop()
        for f in self.null:
            f.close()
        self.tmp.cleanup()

    def run_main(self, *argv, text=None, meta=None):
        with mock.patch.object(coach, "call_llm", return_value=(text or llm_text(37.5), {**META, **(meta or {})})), \
             mock.patch.object(sys, "argv", ["coach.py", "--date", "2026-10-08", *argv]):
            coach.main()
        logs = glob.glob(os.path.join(coach.LOG_DIR, "*_2026-10-08.json"))
        self.assertEqual(len(logs), 1)
        with open(logs[0], encoding="utf-8") as f:
            return json.load(f)

    def test_dry_run_log_has_everything(self):
        log = self.run_main("--dry-run", "--note", "허리 뻐근", text=llm_text(41.25))
        self.assertEqual(log["args"]["note"], "허리 뻐근")
        self.assertEqual(log["context"]["target_date"], "2026-10-08")          # 같은 입력으로 다시 돌릴 수 있게
        self.assertIn("41.25", log["raw_output"])                             # LLM 원문 그대로
        self.assertEqual(log["llm_output"]["plan"]["items"][0]["kg"], 41.25)  # 고치기 전
        self.assertEqual(log["final_plan"]["items"][0]["kg"], 37.5)           # 고친 뒤 (G7)
        self.assertEqual([f["rule"] for f in log["fixes"]], ["G7"])
        self.assertEqual(log["meta"]["cost_usd"], 0.046)
        self.assertIsNone(log["saved"])
        self.assertFalse(os.path.exists(coach.PLAN))

    def test_save_logged(self):
        with mock.patch.object(coach, "save_plan", return_value=("추가", None)) as sp:
            log = self.run_main()
        sp.assert_called_once()
        self.assertEqual(log["saved"], "추가")

    def test_incomplete_response_still_logged(self):
        with self.assertRaises(SystemExit):
            self.run_main("--dry-run", text='{"plan": {"tit', meta={"stop_reason": "max_tokens"})
        with open(glob.glob(os.path.join(coach.LOG_DIR, "*.json"))[0], encoding="utf-8") as f:
            log = json.load(f)
        self.assertIn("max_tokens", log["error"])
        self.assertEqual(log["raw_output"], '{"plan": {"tit')


if __name__ == "__main__":
    unittest.main()
