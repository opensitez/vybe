#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/multiplicative_before_additive
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((1 + 2 * 3)) -eq 7 ] || fail "got $((1 + 2 * 3))"
[ $((10 - 6 / 2)) -eq 7 ] || fail "got $((10 - 6 / 2))"
[ $((7 % 4 + 1)) -eq 4 ] || fail "got $((7 % 4 + 1))"
echo PASS
exit 0
