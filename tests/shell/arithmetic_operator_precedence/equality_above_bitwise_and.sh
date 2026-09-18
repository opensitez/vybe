#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/equality_above_bitwise_and
# 6 & 2 == 2 is 6 & (2 == 2), a classic C precedence trap kept by bash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((6 & 2 == 2)) -eq 0 ] || fail "got $((6 & 2 == 2))"
[ $(((6 & 2) == 2)) -eq 1 ] || fail "grouped got $(((6 & 2) == 2))"
echo PASS
exit 0
