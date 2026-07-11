import unittest

from crawler import split_date_range_by_month


class DateRangeTest(unittest.TestCase):
    def test_splits_cross_month_range(self):
        self.assertEqual(
            split_date_range_by_month("2026-07-13", "2026-08-31"),
            [
                ("2026-07-13", "2026-07-31"),
                ("2026-08-01", "2026-08-31"),
            ],
        )

    def test_keeps_single_month_range(self):
        self.assertEqual(
            split_date_range_by_month("2026-07-13", "2026-07-31"),
            [("2026-07-13", "2026-07-31")],
        )


if __name__ == "__main__":
    unittest.main()
