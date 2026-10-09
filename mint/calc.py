"""Work-order calculations, ported one-to-one from the Excel formulas.

Each function names the Excel column it reproduces (row 11 shown):

  J  Gross profit   =IF(OR(A="",COUNT(H)=0,COUNT(Z)=0),"",ROUND(Z-H,2))
  K  Gross margin   =IF(COUNT(J)=0,"",IF(Z=0,"n.a.",J/Z))
  L  Quote          =IF(OR(A="",COUNT(H)=0,COUNT(Y)=0),"",ROUND(H*(1+Y),2))
  X  Payment due    =IF(OR(A="",COUNT(U)=0,COUNT(AI)=0),"",U+AI)
  AA Actual markup  =IF(COUNT(J)=0,"",IF(H=0,"n.a.",J/H))
  AE Net profit     =IF(OR(COUNT(J)=0,COUNT(AB:AD)<3),"",ROUND(J-SUM(AB:AD),2))
  AF Net margin     =IF(COUNT(AE)=0,"",IF(Z=0,"n.a.",AE/Z))
  AH Balance due    =IF(OR(A="",COUNT(Z)=0,COUNT(AG)=0),"",ROUND(Z-AG,2))
  AJ NTE check      see nte_check()
  AK Input status   see input_status()

A blank Excel cell ("") is None here; "n.a." is kept as the string NA.
"""
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP

NA = "n.a."


def xround(value, digits=2):
    """Excel ROUND: half away from zero (Python's round() is half-to-even).

    Excel works with 15 significant digits, so 494.53499999999997 (the float
    result of 329.69 * 1.5) is 494.535 to Excel and rounds up to 494.54.
    """
    q = Decimal(1).scaleb(-digits)
    d = Decimal(format(value, ".15g")).quantize(q, rounding=ROUND_HALF_UP)
    return float(d)


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _ratio(numerator, denominator):
    if not _num(numerator):
        return None
    if not denominator:
        return NA
    return numerator / denominator


def nte_check(wo, nte, revenue, quote):
    if not wo:
        return None
    if not _num(nte):
        return "Enter approved NTE"
    if _num(revenue):
        return "Invoice exceeds NTE" if revenue > nte else "Invoice within NTE"
    if not _num(quote):
        return "Enter quote inputs"
    return "Quote exceeds NTE" if quote > nte else "Quote within NTE"


def input_status(wo, cost, revenue, expenses):
    if not wo:
        return None
    if not _num(cost):
        return "Enter direct cost"
    if not _num(revenue):
        return "Enter invoice revenue"
    if not all(_num(e) for e in expenses):
        return "Enter expenses (0 if none)"
    return "Complete"


def compute(r, today=None):
    """Return the computed columns for one work order (a dict of inputs)."""
    wo = (r.get("wo_number") or "").strip()
    h, i, y, z = (r.get(k) for k in ("direct_cost", "approved_nte",
                                     "target_markup", "invoice_revenue"))
    expenses = [r.get(k) for k in ("overhead", "admin_fees", "other_expenses")]
    ag, ai = r.get("amount_received"), r.get("payment_terms")
    u = r.get("invoice_sent_date")

    out = {}
    j = xround(z - h) if wo and _num(h) and _num(z) else None
    out["gross_profit"] = j
    out["gross_margin"] = _ratio(j, z)
    out["quote"] = xround(h * (1 + y)) if wo and _num(h) and _num(y) else None
    out["actual_markup"] = _ratio(j, h)
    ae = (xround(j - sum(expenses))
          if _num(j) and all(_num(e) for e in expenses) else None)
    out["net_profit"] = ae
    out["net_margin"] = _ratio(ae, z)
    out["balance_due"] = xround(z - ag) if wo and _num(z) and _num(ag) else None

    due = None
    if wo and u and _num(ai):
        d = u if isinstance(u, date) else date.fromisoformat(str(u))
        due = (d + timedelta(days=int(ai))).isoformat()
    out["payment_due"] = due

    out["nte_check"] = nte_check(wo, i, z, out["quote"])
    out["input_status"] = input_status(wo, h, z, expenses)

    # Extras the web app adds on top of the spreadsheet (not Excel columns).
    out.update(receivable_status(r, out, today or date.today()))
    return out


AGING_BUCKETS = ["Not yet due", "1-30 days", "31-60 days", "61-90 days",
                 "Over 90 days", "No due date"]


def receivable_status(r, computed, today):
    """Accounts-receivable view of a work order: is money still owed, how late."""
    bal = computed["balance_due"]
    owed = (_num(bal) and bal > 0.004
            and r.get("invoice_status") not in ("Void", "Not invoiced", "Draft"))
    if not owed:
        return {"owed": False, "days_overdue": None, "aging_bucket": None,
                "overdue": False}
    due = computed["payment_due"]
    if not due:
        return {"owed": True, "days_overdue": None, "aging_bucket": "No due date",
                "overdue": False}
    days = (today - date.fromisoformat(due)).days
    if days <= 0:
        bucket = "Not yet due"
    elif days <= 30:
        bucket = "1-30 days"
    elif days <= 60:
        bucket = "31-60 days"
    elif days <= 90:
        bucket = "61-90 days"
    else:
        bucket = "Over 90 days"
    return {"owed": True, "days_overdue": max(days, 0), "aging_bucket": bucket,
            "overdue": days > 0}


def totals(rows):
    """The summary block at the top of the sheet (A5:F5), plus app extras.

    rows are dicts that already contain the computed columns.
    """
    def total(key):
        return xround(sum(r[key] for r in rows if _num(r.get(key))))

    return {
        "work_orders": sum(1 for r in rows if r.get("wo_number")),          # A5
        "invoiced_revenue": total("invoice_revenue"),                       # B5
        "gross_profit": total("gross_profit"),                              # C5
        "net_profit": total("net_profit"),                                  # D5
        "balance_due": total("balance_due"),                                # E5
        "needs_input": sum(1 for r in rows                                  # F5
                           if str(r.get("input_status") or "").startswith("Enter")),
        "direct_cost": total("direct_cost"),
        "amount_received": total("amount_received"),
        "overdue_amount": xround(sum(r["balance_due"] for r in rows if r.get("overdue"))),
        "overdue_count": sum(1 for r in rows if r.get("overdue")),
        "nte_exceptions": sum(1 for r in rows
                              if str(r.get("nte_check") or "").endswith("exceeds NTE")),
    }
