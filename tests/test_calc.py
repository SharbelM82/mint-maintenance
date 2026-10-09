from datetime import date

from mint.calc import NA, compute, totals, xround

FULL = {"wo_number": "1", "direct_cost": 1000.0, "target_markup": 0.4,
        "invoice_revenue": 1400.0, "overhead": 100.0, "admin_fees": 30.0,
        "other_expenses": 20.0, "approved_nte": 1500.0, "amount_received": 400.0,
        "invoice_sent_date": "2026-01-10", "payment_terms": 30,
        "invoice_status": "Sent"}


def test_worked_example_from_guide_sheet():
    out = compute(FULL, today=date(2026, 1, 15))
    assert out["quote"] == 1400.0                      # L
    assert out["gross_profit"] == 400.0                # J
    assert out["gross_margin"] == 400 / 1400           # K
    assert out["actual_markup"] == 0.4                 # AA
    assert out["net_profit"] == 250.0                  # AE
    assert out["net_margin"] == 250 / 1400             # AF
    assert out["balance_due"] == 1000.0                # AH
    assert out["payment_due"] == "2026-02-09"          # X = U + AI
    assert out["nte_check"] == "Invoice within NTE"    # AJ
    assert out["input_status"] == "Complete"           # AK


def test_blank_wo_number_blanks_everything_like_excel():
    out = compute({**FULL, "wo_number": ""})
    for k in ("gross_profit", "quote", "payment_due", "balance_due", "nte_check",
              "input_status", "net_profit"):
        assert out[k] is None, k


def test_missing_inputs_leave_results_blank():
    out = compute({"wo_number": "7", "direct_cost": 500.0})
    assert out["gross_profit"] is None and out["gross_margin"] is None
    assert out["quote"] is None
    assert out["input_status"] == "Enter invoice revenue"
    assert out["nte_check"] == "Enter approved NTE"
    out = compute({"wo_number": "7"})
    assert out["input_status"] == "Enter direct cost"


def test_expenses_need_all_three_values_zero_counts():
    row = {**FULL, "other_expenses": None}
    assert compute(row)["net_profit"] is None
    assert compute(row)["input_status"] == "Enter expenses (0 if none)"
    row["other_expenses"] = 0.0
    assert compute(row)["net_profit"] == 270.0


def test_zero_denominators_show_na():
    out = compute({**FULL, "direct_cost": 0.0, "invoice_revenue": 0.0})
    assert out["gross_profit"] == 0.0
    assert out["gross_margin"] == NA
    assert out["actual_markup"] == NA
    assert out["net_margin"] == NA


def test_nte_check_branches():
    base = {"wo_number": "1", "approved_nte": 1000.0}
    assert compute({**base})["nte_check"] == "Enter quote inputs"
    assert compute({**base, "direct_cost": 800.0, "target_markup": 0.3})["nte_check"] == "Quote exceeds NTE"
    assert compute({**base, "direct_cost": 700.0, "target_markup": 0.3})["nte_check"] == "Quote within NTE"
    # Invoice revenue takes priority over the quote.
    assert compute({**base, "direct_cost": 900.0, "target_markup": 0.5,
                    "invoice_revenue": 1000.0})["nte_check"] == "Invoice within NTE"
    assert compute({**base, "invoice_revenue": 1000.01})["nte_check"] == "Invoice exceeds NTE"


def test_excel_rounding_is_half_away_from_zero():
    assert xround(2.675) == 2.68
    assert xround(-2.675) == -2.68
    assert compute({"wo_number": "1", "direct_cost": 10.05, "target_markup": 0.5})["quote"] == 15.08
    # Float noise: 329.69 * 1.5 = 494.53499999999997, Excel shows 494.54.
    assert compute({"wo_number": "1", "direct_cost": 329.69, "target_markup": 0.5})["quote"] == 494.54


def test_receivables_aging():
    out = compute(FULL, today=date(2026, 3, 1))     # due 2026-02-09
    assert out["overdue"] and out["days_overdue"] == 20 and out["aging_bucket"] == "1-30 days"
    out = compute(FULL, today=date(2026, 2, 9))
    assert not out["overdue"] and out["aging_bucket"] == "Not yet due"
    out = compute({**FULL, "invoice_status": "Void"}, today=date(2026, 6, 1))
    assert not out["owed"]
    out = compute({**FULL, "amount_received": 1400.0})
    assert not out["owed"]


def test_totals_match_sheet_summary():
    rows = []
    for r in (FULL, {"wo_number": "2", "direct_cost": 100.0}, {"wo_number": ""}):
        rows.append({**r, **compute(r, today=date(2026, 1, 15))})
    t = totals(rows)
    assert t["work_orders"] == 2
    assert t["invoiced_revenue"] == 1400.0
    assert t["gross_profit"] == 400.0
    assert t["net_profit"] == 250.0
    assert t["balance_due"] == 1000.0
    assert t["needs_input"] == 1
