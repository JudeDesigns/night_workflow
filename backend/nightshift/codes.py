"""Code normalization and matching.

Spec rules (§2):
  * Trim whitespace.
  * A number equals its text form (123 == "123").
  * Shopping History exception: a trailing `U` or `C` may be present on one
    side. Treat codes as equal when the only difference is one trailing letter
    (U or C) on either side.

`norm_code` returns the canonical text form. `code_match_loose` is the fuzzy
variant used when one side of the comparison is from Shopping History.
"""
from __future__ import annotations

from typing import Any


def norm_code(value: Any) -> str:
    """Normalize a code value to its canonical text form.

    - None / empty string -> ""
    - Numbers: drop trailing `.0` introduced by Excel float reads
              (e.g. 1600053162.0 -> "1600053162").
    - Strings: stripped.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        # Excel never stores codes as booleans, but be defensive.
        return ""
    if isinstance(value, (int,)):
        return str(value).strip()
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value)).strip()
        return repr(value).strip()
    s = str(value).strip()
    # Strip a trailing ".0" that comes from Excel reading integer-valued
    # numeric cells as floats and then someone str()ing them upstream.
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def _strip_trailing_uc(code: str) -> str:
    if len(code) > 1 and code[-1] in ("U", "C", "u", "c"):
        return code[:-1]
    return code


def code_match_loose(a: Any, b: Any) -> bool:
    """Return True if `a` and `b` are equal under the Shopping History rule.

    Either side may carry an extra trailing U or C; otherwise the canonical
    text forms must match exactly (case-insensitively).
    """
    na = norm_code(a).upper()
    nb = norm_code(b).upper()
    if not na or not nb:
        return False
    if na == nb:
        return True
    return _strip_trailing_uc(na) == _strip_trailing_uc(nb)


def code_key(value: Any) -> str:
    """Hashable key for strict matches (item list, Inventory)."""
    return norm_code(value).upper()


def code_key_loose(value: Any) -> str:
    """Hashable key for Shopping-History-style fuzzy matches.

    Two values map to the same key iff `code_match_loose` would return True.
    Implementation: strip trailing U/C unconditionally before keying.
    """
    return _strip_trailing_uc(norm_code(value).upper())
