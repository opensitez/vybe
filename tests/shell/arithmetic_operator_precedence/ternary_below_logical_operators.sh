#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/ternary_below_logical_operators
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 || 0 ? 8 : 9)) -eq 8 ] || fail "got $((1 || 0 ? 8 : 9))"
[ $((0 && 1 ? 8 : 9)) -eq 9 ] || fail "got $((0 && 1 ? 8 : 9))"
[ $((2 > 1 ? 10 + 1 : 0)) -eq 11 ] || fail "arithmetic inside branch got $((2 > 1 ? 10 + 1 : 0))"
echo PASS
exit 0
