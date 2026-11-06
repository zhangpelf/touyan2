"""
tests/test_builder.py
"""

import os
import unittest
import tempfile
from touyan2.models import ReportContent
from touyan2.builder import generate_report


class TestReportBuilder(unittest.TestCase):
    def test_generate_report_basic(self):
        content = ReportContent("单元测试智能制造")
        content.ch1_highlights = ["**测试亮点**：各项指标达标"]
        content.ch1_basic_info_rows = [["项目主体", "单元测试公司"]]
        content.ch7_score_rows = [["综合得分", "100%", "95", "建议推进"]]

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            out = generate_report(content, tmp_path)
            self.assertTrue(os.path.exists(out))
            self.assertGreater(os.path.getsize(out), 1000)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


if __name__ == "__main__":
    unittest.main()
