# -*- coding: utf-8 -*-
"""plan.json 병합 저장 테스트 (임시 폴더에서). 실행: python -m unittest discover -s tests -v"""
import glob, json, os, sys, tempfile, unittest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import coach


def plan(pid, date, title="t"):
    return {"id": pid, "date": date, "title": title, "note": "", "items": [{"ex": "레그컬", "kg": 40, "reps": 12, "sets": 3}]}


class TestMerge(unittest.TestCase):
    def test_replace_same_id_keep_others(self):
        doc = {"plans": [plan("2026-09-25", "2026-09-25"), plan("ai-2026-10-08", "2026-10-08", "old")]}
        new, action = coach.merge_plans(doc, plan("ai-2026-10-08", "2026-10-08", "new"))
        self.assertEqual(action, "교체")
        self.assertEqual([p["id"] for p in new["plans"]], ["2026-09-25", "ai-2026-10-08"])   # 자리 유지
        self.assertEqual(new["plans"][1]["title"], "new")
        self.assertEqual(doc["plans"][1]["title"], "old")                                    # 원본은 안 바뀜

    def test_append_new_id(self):
        new, action = coach.merge_plans({"plans": [plan("2026-09-25", "2026-09-25")]}, plan("ai-2026-10-10", "2026-10-10"))
        self.assertEqual(action, "추가")
        self.assertEqual(len(new["plans"]), 2)

    def test_empty_doc(self):
        new, action = coach.merge_plans(None, plan("ai-2026-10-10", "2026-10-10"))
        self.assertEqual((action, len(new["plans"])), ("추가", 1))


class TestSave(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "plan.json")
        self.bak = os.path.join(self.tmp.name, "백업")

    def tearDown(self):
        self.tmp.cleanup()

    def read(self):
        with open(self.path, encoding="utf-8") as f:
            return json.load(f)

    def test_new_file_no_backup(self):
        action, bak = coach.save_plan(plan("ai-2026-10-08", "2026-10-08"), self.path, self.bak)
        self.assertEqual((action, bak), ("추가", None))
        self.assertEqual(self.read()["plans"][0]["id"], "ai-2026-10-08")

    def test_backup_then_replace(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump({"plans": [plan("2026-09-25", "2026-09-25")]}, f)
        action, bak = coach.save_plan(plan("ai-2026-10-08", "2026-10-08"), self.path, self.bak)
        self.assertTrue(os.path.exists(bak))
        with open(bak, encoding="utf-8") as f:
            self.assertEqual(len(json.load(f)["plans"]), 1)                 # 백업 = 쓰기 전 내용
        self.assertEqual([p["id"] for p in self.read()["plans"]], ["2026-09-25", "ai-2026-10-08"])
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_backups_pruned(self):
        os.makedirs(self.bak)
        for i in range(12):
            open(os.path.join(self.bak, f"plan_20260101-0000{i:02d}.json"), "w").close()
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump({"plans": []}, f)
        coach.save_plan(plan("ai-2026-10-08", "2026-10-08"), self.path, self.bak)
        self.assertEqual(len(glob.glob(os.path.join(self.bak, "plan_*.json"))), coach.PLAN_BACKUPS)

    def test_broken_file_not_overwritten(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("{깨진")
        with self.assertRaises(json.JSONDecodeError):
            coach.save_plan(plan("ai-2026-10-08", "2026-10-08"), self.path, self.bak)
        with open(self.path, encoding="utf-8") as f:
            self.assertEqual(f.read(), "{깨진")


class TestDelete(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "plan.json")
        self.bak = os.path.join(self.tmp.name, "백업")
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump({"plans": [plan("2026-09-25", "2026-09-25"), plan("ai-2026-10-08", "2026-10-08", "AI")]}, f)

    def tearDown(self):
        self.tmp.cleanup()

    def test_delete_ai_plan_with_backup(self):
        gone, bak = coach.delete_plan("ai-2026-10-08", self.path, self.bak)
        self.assertEqual(gone["title"], "AI")
        self.assertEqual([p["id"] for p in coach.list_plans(self.path)], ["2026-09-25"])   # 손으로 쓴 계획은 남음
        with open(bak, encoding="utf-8") as f:
            self.assertEqual(len(json.load(f)["plans"]), 2)                                 # 백업 = 지우기 전

    def test_missing_id(self):
        self.assertEqual(coach.delete_plan("ai-2026-12-31", self.path, self.bak), (None, None))
        self.assertFalse(os.path.exists(self.bak))                                          # 바꾼 게 없으면 백업도 안 만듦

    def test_manual_plan_protected(self):
        with self.assertRaises(coach.CoachError):
            coach.delete_plan("2026-09-25", self.path, self.bak)


if __name__ == "__main__":
    unittest.main()
