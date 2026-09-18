#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/unary_not_and_complement_above_binary
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((!0 + 1)) -eq 2 ] || fail "got $((!0 + 1))"
[ $((~1 + 1)) -eq -1 ] || fail "got $((~1 + 1))"
[ $((!(0 + 1))) -eq 0 ] || fail "grouped got $((!(0 + 1)))"
echo PASS
exit 0
