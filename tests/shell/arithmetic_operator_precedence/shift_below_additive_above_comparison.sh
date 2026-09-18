#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/shift_below_additive_above_comparison
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 + 2 << 1)) -eq 6 ] || fail "1 + 2 << 1 want 6 got $((1 + 2 << 1))"
[ $((1 + (2 << 1))) -eq 5 ] || fail "grouped want 5"
[ $((1 << 2 < 5)) -eq 1 ] || fail "1 << 2 < 5 want 1 got $((1 << 2 < 5))"
echo PASS
exit 0
