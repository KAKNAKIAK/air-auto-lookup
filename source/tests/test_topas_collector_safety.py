from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import topas_live_collector as collector_mod


class FakeDriver:
    def __init__(self):
        self.quit_count = 0

    def quit(self):
        self.quit_count += 1


class ReusableCollector(collector_mod.TopasLiveCollector):
    def __init__(self):
        super().__init__(["127.0.0.1:9222"], timeout=1.0)
        self.connect_count = 0
        self.switch_count = 0
        self.fake_driver = FakeDriver()

    def connect(self):
        self.connect_count += 1
        return self.fake_driver, "127.0.0.1:9222"

    def switch_to_topas(self, driver):
        self.switch_count += 1


class TopasCollectorSafetyTests(unittest.TestCase):
    def test_target_looks_like_topas(self):
        self.assertTrue(collector_mod.target_looks_like_topas({"url": "https://www.topassellconnect.com/", "title": ""}))
        self.assertTrue(collector_mod.target_looks_like_topas({"url": "about:blank", "title": "TOPAS Sell Connect"}))
        self.assertFalse(collector_mod.target_looks_like_topas({"url": "https://example.com/", "title": "ERP"}))

    def test_require_topas_debug_target_rejects_non_topas_browser(self):
        with patch.object(collector_mod, "debugger_targets", return_value=[{"url": "https://example.com/", "title": "ERP"}]):
            with self.assertRaises(RuntimeError):
                collector_mod.require_topas_debug_target("127.0.0.1:9223")

    def test_collect_reuses_single_driver_until_close(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            log_dir = Path(temp_dir)
            collector = ReusableCollector()
            self.assertEqual(collector.collect([], log_dir), [])
            self.assertEqual(collector.collect([], log_dir), [])
            self.assertEqual(collector.connect_count, 1)
            self.assertEqual(collector.fake_driver.quit_count, 0)
            collector.close()
            self.assertEqual(collector.fake_driver.quit_count, 1)


if __name__ == "__main__":
    unittest.main()
