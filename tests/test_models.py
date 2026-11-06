"""
tests/test_models.py
"""

import os
import unittest
import tempfile
from touyan2.models import ReportContent


class TestReportModels(unittest.TestCase):
    def test_report_content_defaults(self):
        content = ReportContent("测试科技")
        self.assertEqual(content.company, "测试科技")
        self.assertEqual(content.report_title, "投研分析报告")
        self.assertEqual(len(content.ch1_basic_info_headers), 2)

    def test_report_content_dict_and_json(self):
        content = ReportContent("测试科技")
        content.ch1_highlights.append("亮点一")
        content.ch1_basic_info_rows.append(["企业全称", "测试科技有限公司"])

        d = content.to_dict()
        self.assertEqual(d["company"], "测试科技")
        self.assertEqual(d["ch1_highlights"], ["亮点一"])

        rebuilt = ReportContent.from_dict(d)
        self.assertEqual(rebuilt.company, "测试科技")
        self.assertEqual(rebuilt.ch1_highlights, ["亮点一"])
        self.assertEqual(rebuilt.ch1_basic_info_rows[0][1], "测试科技有限公司")

        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            content.save_json(tmp_path)
            loaded = ReportContent.load_json(tmp_path)
            self.assertEqual(loaded.company, "测试科技")
            self.assertEqual(loaded.ch1_highlights, ["亮点一"])
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


if __name__ == "__main__":
    unittest.main()
