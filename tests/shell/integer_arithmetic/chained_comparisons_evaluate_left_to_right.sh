#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/chained_comparisons_evaluate_left_to_right
# 5 > 3 > 1 is (5 > 3) > 1, i.e. 1 > 1, which is 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((5 > 3 > 1)) -eq 0 ] || fail "got $((5 > 3 > 1))"
[ $((1 < 2 < 3)) -eq 1 ] || fail "got $((1 < 2 < 3))"
echo PASS
exit 0
