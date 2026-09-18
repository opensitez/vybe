#!/usr/bin/env bash
# vybe-test: bash/integer_arithmetic/integer_modulo_sign_follows_dividend
# Remainder keeps the sign of the dividend in Bash/C arithmetic.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((-7 % 3)) -eq -1 ] || fail "-7%3 got $((-7 % 3))"
[ $((7 % -3)) -eq 1 ] || fail "7%-3 got $((7 % -3))"
[ $((-7 % -3)) -eq -1 ] || fail "-7%-3 got $((-7 % -3))"
echo PASS
exit 0
