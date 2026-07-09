import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from engine.amount_words import amount_to_words


# 方案 §3.3 的 10 条必测边界用例
CASES = [
    (898, "EIGHT HUNDRED NINETY EIGHT ONLY"),
    (100, "ONE HUNDRED ONLY"),
    (1000, "ONE THOUSAND ONLY"),
    (1005, "ONE THOUSAND FIVE ONLY"),
    (21015, "TWENTY ONE THOUSAND FIFTEEN ONLY"),
    (100000, "ONE HUNDRED THOUSAND ONLY"),
    (1234567, "ONE MILLION TWO HUNDRED THIRTY FOUR THOUSAND FIVE HUNDRED SIXTY SEVEN ONLY"),
    (100.5, "ONE HUNDRED AND CENTS 50/100 ONLY"),
    (0.99, "ZERO AND CENTS 99/100 ONLY"),
    (898.05, "EIGHT HUNDRED NINETY EIGHT AND CENTS 05/100 ONLY"),
]


@pytest.mark.parametrize("value,expected", CASES)
def test_plan_cases(value, expected):
    assert amount_to_words(value) == expected


def test_float_edge_rounds_up():
    assert amount_to_words(1.999999) == "TWO ONLY"


def test_negative_rejected():
    with pytest.raises(ValueError):
        amount_to_words(-1)
