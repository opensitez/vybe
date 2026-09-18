#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_negative_division_rounds_toward_zero
# Bash integer division truncates toward zero.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((-7 / 3)) -eq -2 ] || fail "-7/3 got $((-7 / 3))"
[ $((7 / -3)) -eq -2 ] || fail "7/-3 got $((7 / -3))"
[ $((-7 / -3)) -eq 2 ] || fail "-7/-3 got $((-7 / -3))"
echo PASS
exit 0
