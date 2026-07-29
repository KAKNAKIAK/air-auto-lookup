from pathlib import Path
import json
import sys
import tempfile
import tkinter as tk
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import air_auto_lookup_mvp as app_mod


class AirAutoLookupGuiTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = app_mod.AirAutoLookupApp(self.root)

    def tearDown(self):
        self.root.destroy()

    def test_new_and_resume_actions_are_separate_on_first_screen(self):
        self.assertEqual(self.app.primary_run_button.cget("text"), "▶ 새 실행")
        self.assertEqual(self.app.resume_run_button.cget("text"), "↻ 이어 실행")
        self.assertEqual(str(self.app.resume_run_button.cget("state")), str(tk.DISABLED))

    def test_new_run_button_starts_new_run_even_when_save_is_selected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            Path(temp_dir, "run.json").write_text("{}", encoding="utf-8")
            self.app.current_log_dir.set(temp_dir)
            self.app.update_primary_run_button()
            self.app.auto_run_all = Mock()
            self.app.collect_topas = Mock()

            self.app.primary_run_button.invoke()

            self.app.auto_run_all.assert_called_once_with()
            self.app.collect_topas.assert_not_called()

    def test_resume_button_keeps_mismatch_guard_and_points_to_new_run(self):
        selected_masters = self.app.selected_masters()
        run_doc = {
            "startDate": "2000-01-01",
            "endDate": self.app.end_date.get(),
            "productDays": int(self.app.product_days.get()),
            "selectedRoutes": sorted({master.route for master in selected_masters}),
            "selectedMasterKeys": [master.key for master in selected_masters],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            Path(temp_dir, "run.json").write_text(
                json.dumps(run_doc, ensure_ascii=False),
                encoding="utf-8",
            )
            self.app.current_log_dir.set(temp_dir)
            self.app.update_primary_run_button()

            with patch.object(app_mod.messagebox, "showerror") as showerror:
                self.app.resume_run_button.invoke()

            showerror.assert_called_once()
            title, message = showerror.call_args.args
            self.assertEqual(title, "수집 조건 확인")
            self.assertIn("'새 실행'", message)

    def test_excel_open_button_is_disabled_until_result_exists(self):
        self.assertEqual(self.app.excel_open_button.cget("text"), "엑셀 열기")
        self.assertEqual(str(self.app.excel_open_button.cget("state")), str(tk.DISABLED))

    def test_excel_open_button_activates_and_opens_completed_result(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            result_path = Path(temp_dir, "result.xlsx")
            result_path.write_bytes(b"test workbook placeholder")
            Path(temp_dir, "run.json").write_text(
                json.dumps({"calculatedExcel": str(result_path)}, ensure_ascii=False),
                encoding="utf-8",
            )
            self.app.current_log_dir.set(temp_dir)
            self.app.update_primary_run_button()

            self.assertEqual(str(self.app.excel_open_button.cget("state")), str(tk.NORMAL))
            with patch.object(app_mod.os, "startfile") as startfile:
                self.app.excel_open_button.invoke()
            startfile.assert_called_once_with(result_path)


if __name__ == "__main__":
    unittest.main()
