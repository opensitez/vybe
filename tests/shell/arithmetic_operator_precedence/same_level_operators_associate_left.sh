#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/same_level_operators_associate_left
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((10 - 4 - 3)) -eq 3 ] || fail "got $((10 - 4 - 3))"
[ $((100 / 10 / 2)) -eq 5 ] || fail "got $((100 / 10 / 2))"
[ $((2 ** 3 ** 2)) -eq 512 ] || fail "exponent is the right-associative exception, got $((2 ** 3 ** 2))"
echo PASS
exit 0
