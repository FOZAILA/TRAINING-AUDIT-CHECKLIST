"""Small helpers shared by all modules."""
import datetime as dt
import re

_DATE_FORMATS = ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%d-%b-%y", "%d/%m/%y")
_EMPTY_WORDS = {"", "-", "--", "---", "NA", "N/A", "NIL", "NONE", "NULL", "PENDING"}


def clean(text):
    """Collapse whitespace/newlines of a header or cell text."""
    return re.sub(r"\s+", " ", str(text)).strip() if text is not None else ""


def norm(text):
    """Lower-case letters+digits only, used for fuzzy header matching."""
    return re.sub(r"[^a-z0-9]+", " ", clean(text).lower()).strip()


def to_date(value):
    """Return a date, True (a value exists but is not a readable date) or None (empty)."""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = clean(value)
    if text.upper() in _EMPTY_WORDS:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return True


def status_of(value, ttype, as_of, soon_days):
    """none | valid | soon | expired.  Only 'expiry' trainings can expire."""
    if value is None:
        return "none"
    if ttype != "expiry" or value is True:
        return "valid"
    if value < as_of:
        return "expired"
    if value <= as_of + dt.timedelta(days=soon_days):
        return "soon"
    return "valid"


def pct(part, whole):
    return round(part * 100.0 / whole, 1) if whole else None
