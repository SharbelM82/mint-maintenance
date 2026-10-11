"""Parsing user/Excel input into stored values, and formatting for display."""
from datetime import date, datetime

from .calc import NA


class InvalidValue(ValueError):
    pass


def _to_number(raw, field):
    s = str(raw).strip().replace("$", "").replace(",", "").replace(" ", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    pct = s.endswith("%")
    s = s.rstrip("%")
    try:
        n = float(s)
    except ValueError:
        raise InvalidValue(f"{field.label}: '{raw}' is not a number")
    return (-n if neg else n), pct


def parse(field, raw, source="form"):
    """Convert a raw value to the stored form. Blank becomes None.

    source="form": percents are typed as 40 meaning 40%.
    source="excel": percents are cell values (0.4) unless written as "40%".
    """
    if raw is None:
        return None
    if isinstance(raw, str):
        raw = raw.strip()
        if raw == "":
            return None
    kind = field.kind

    if kind in ("text", "textarea"):
        if isinstance(raw, float) and raw.is_integer():
            raw = int(raw)      # WO # 1234 read from Excel as 1234.0
        return str(raw)

    if kind == "choice":
        for c in field.choices:
            if str(raw).strip().lower() == c.lower():
                return c
        raise InvalidValue(f"{field.label}: '{raw}' is not one of "
                           + ", ".join(field.choices))

    if kind == "date":
        if isinstance(raw, datetime):
            return raw.date().isoformat()
        if isinstance(raw, date):
            return raw.isoformat()
        s = str(raw).strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(s, fmt).date().isoformat()
            except ValueError:
                pass
        raise InvalidValue(f"{field.label}: '{raw}' is not a date (use mm/dd/yyyy)")

    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        n, pct_sign = float(raw), False
    else:
        n, pct_sign = _to_number(raw, field)

    if kind == "money":
        if n < 0:
            raise InvalidValue(f"{field.label} cannot be negative")
        return round(n, 2)
    if kind == "int":
        if n < 0 or not float(n).is_integer():
            raise InvalidValue(f"{field.label} must be a whole number of 0 or more")
        return int(n)
    if kind == "percent":
        if source == "form" or pct_sign:
            n = n / 100
        if n < -1:
            raise InvalidValue(f"{field.label} cannot be below -100%")
        return round(n, 6)
    raise InvalidValue(f"Unknown field type {kind}")


def form_value(field, value):
    """Value to pre-fill in an <input> (percent shown as 40, not 0.4)."""
    if value is None:
        return ""
    if field.kind == "percent" and isinstance(value, (int, float)):
        return f"{value * 100:g}"
    if field.kind == "money" and isinstance(value, (int, float)):
        return f"{value:.2f}"
    return value


def display(field, value):
    """Human-readable value for tables and detail pages."""
    if value is None or value == "":
        return ""
    if value == NA:
        return NA
    kind = field.kind
    if kind == "money":
        return f"(${-value:,.2f})" if value < 0 else f"${value:,.2f}"
    if kind == "percent":
        return f"{value * 100:.1f}%"
    if kind == "date":
        try:
            return date.fromisoformat(value).strftime("%m/%d/%y")
        except (TypeError, ValueError):
            return value
    return value
