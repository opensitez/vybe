#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/comparison_below_additive
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 + 1 == 2)) -eq 1 ] || fail "got $((1 + 1 == 2))"
[ $((1 < 2 + 3)) -eq 1 ] || fail "got $((1 < 2 + 3))"
[ $((2 * 3 >= 6)) -eq 1 ] || fail "got $((2 * 3 >= 6))"
echo PASS
exit 0
