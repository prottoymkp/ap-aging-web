import datetime as dt
import unittest

import pandas as pd

from engine import fifo_aging


class TestFifoAging(unittest.TestCase):
    def test_balanced_with_future_dated_row_results_in_zero_outstanding(self):
        as_of = dt.date(2024, 1, 31)
        tx_df = pd.DataFrame(
            [
                {"date": dt.date(2024, 1, 10), "credit": 100.0, "debit": 0.0, "net": 100.0},
                {"date": dt.date(2024, 1, 15), "credit": 0.0, "debit": 100.0, "net": -100.0},
                # Future-dated row where credit and debit offset (net=0), so no unpaid liability remains.
                {"date": dt.date(2024, 2, 10), "credit": 50.0, "debit": 50.0, "net": 0.0},
            ]
        )

        total_payable, total_paid, balance, _, _, bs, _ = fifo_aging(tx_df, as_of)

        self.assertEqual(total_payable, 150.0)
        self.assertEqual(total_paid, 150.0)
        self.assertEqual(balance, 0.0)
        self.assertEqual(bs["future_dated_unpaid"], 0.0)
        self.assertEqual(bs["advance_overpaid"], 0.0)

    def test_true_overpayment_sets_advance_equal_to_absolute_negative_balance(self):
        as_of = dt.date(2024, 1, 31)
        tx_df = pd.DataFrame(
            [
                {"date": dt.date(2024, 1, 10), "credit": 50.0, "debit": 0.0, "net": 50.0},
                {"date": dt.date(2024, 1, 20), "credit": 0.0, "debit": 80.0, "net": -80.0},
            ]
        )

        _, _, balance, _, _, bs, _ = fifo_aging(tx_df, as_of)

        self.assertLess(balance, 0.0)
        self.assertEqual(bs["advance_overpaid"], abs(balance))
        self.assertEqual(bs["0-30"], 0.0)
        self.assertEqual(bs["31-60"], 0.0)
        self.assertEqual(bs["61-90"], 0.0)
        self.assertEqual(bs["91-180"], 0.0)
        self.assertEqual(bs["181-365"], 0.0)
        self.assertEqual(bs[">365"], 0.0)
        self.assertEqual(bs["future_dated_unpaid"], 0.0)
        self.assertEqual(bs["unknown_date_unpaid"], 0.0)

    def test_future_unpaid_liability_populates_future_bucket_without_advance(self):
        as_of = dt.date(2024, 1, 31)
        tx_df = pd.DataFrame(
            [
                {"date": dt.date(2024, 1, 10), "credit": 20.0, "debit": 0.0, "net": 20.0},
                {"date": dt.date(2024, 2, 5), "credit": 120.0, "debit": 0.0, "net": 120.0},
                {"date": dt.date(2024, 1, 12), "credit": 0.0, "debit": 20.0, "net": -20.0},
            ]
        )

        _, _, balance, _, _, bs, _ = fifo_aging(tx_df, as_of)

        self.assertGreater(balance, 0.0)
        self.assertEqual(bs["future_dated_unpaid"], 120.0)
        self.assertEqual(bs["advance_overpaid"], 0.0)

        bucket_total = (
            bs["0-30"]
            + bs["31-60"]
            + bs["61-90"]
            + bs["91-180"]
            + bs["181-365"]
            + bs[">365"]
            + bs["future_dated_unpaid"]
            + bs["unknown_date_unpaid"]
            - bs["advance_overpaid"]
        )
        self.assertEqual(balance, bucket_total)


if __name__ == "__main__":
    unittest.main()
