"""Merge sub-company names (BUSAN-UIG, CPECC DIV 1-Sigma ...) into one parent company.

The parent list lives in config/settings.json -> company_parents.  The first pattern
that matches wins, so keep specific names (CPECC ZONE 1) above general ones (CPECC).
"""
import re

from .util import clean


def merge_company(raw, settings):
    """Return (parent_name, was_mapped)."""
    name = clean(raw).upper()
    if not name:
        return "(BLANK)", True
    for parent in settings["company_parents"]:
        for pattern in parent["patterns"]:
            if re.search(pattern, name):
                return parent["name"], True
    # Not in the list: keep the part before a dash if it is long enough, and report it.
    m = re.match(r"^(.{4,}?)\s*[-–—]", name)
    return (m.group(1).strip() if m else name), False
