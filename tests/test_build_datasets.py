import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from build_datasets import parse_period_from_path


class ParsePeriodFromPathTest(unittest.TestCase):
    def test_month_name_with_year_in_filename(self):
        path = 'Jampea/2026/LAP. OPS. UPP JAMPEA PERIODE MEI 2026.pdf'
        self.assertEqual(parse_period_from_path(path), (2026, 5))

    def test_numeric_month_prefix(self):
        for bulan in (1, 2, 4):
            path = f'Jampea/2026/2026_{bulan:02d}_730104_JAMPEA.pdf'
            self.assertEqual(parse_period_from_path(path), (2026, bulan))

    def test_year_falls_back_to_parent_directory(self):
        path = 'Jampea/2026/LAPORAN OPERASIONAL JUNI.pdf'
        self.assertEqual(parse_period_from_path(path), (2026, 6))

    def test_unrecognised_month_is_not_defaulted_to_january(self):
        path = 'Jampea/2026/LAPORAN OPERASIONAL.pdf'
        self.assertEqual(parse_period_from_path(path), (2026, None))


if __name__ == '__main__':
    unittest.main()
