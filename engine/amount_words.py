"""
Amount -> English words (uppercase), for Proforma Invoice totals.

Style follows the confirmed sample pi__20260605_XNY-AD260045.pdf:
  898     -> EIGHT HUNDRED NINETY EIGHT ONLY          (no "AND" inside integers)
  898.5   -> EIGHT HUNDRED NINETY EIGHT AND CENTS 50/100 ONLY
  898.05  -> EIGHT HUNDRED NINETY EIGHT AND CENTS 05/100 ONLY
"""

_ONES = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE",
         "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN",
         "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
_TENS = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]
_SCALES = ["", " THOUSAND", " MILLION", " BILLION"]


def _three(n):
    """0-999 -> words (no leading/trailing spaces)."""
    parts = []
    if n >= 100:
        parts.append(_ONES[n // 100] + " HUNDRED")
        n %= 100
    if n >= 20:
        if n % 10:
            parts.append(_TENS[n // 10] + " " + _ONES[n % 10])
        else:
            parts.append(_TENS[n // 10])
    elif n > 0:
        parts.append(_ONES[n])
    return " ".join(parts)


def _integer_words(n):
    if n == 0:
        return "ZERO"
    groups = []
    while n > 0:
        groups.append(n % 1000)
        n //= 1000
    parts = []
    for i in range(len(groups) - 1, -1, -1):
        if groups[i]:
            parts.append(_three(groups[i]) + _SCALES[i])
    return " ".join(parts)


def amount_to_words(amount):
    """Convert a non-negative amount to uppercase English words ending in ONLY."""
    if amount is None:
        raise ValueError("amount is None")
    amount = round(float(amount) + 1e-9, 2)
    if amount < 0:
        raise ValueError("amount must be >= 0")
    integer = int(amount)
    cents = int(round((amount - integer) * 100))
    if cents == 100:  # float edge, e.g. 1.999999
        integer, cents = integer + 1, 0
    words = _integer_words(integer)
    if cents:
        return f"{words} AND CENTS {cents:02d}/100 ONLY"
    return f"{words} ONLY"
