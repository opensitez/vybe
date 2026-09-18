#!/usr/bin/env bash
# vybe-test: bash/arithmetic_operator_precedence/unary_minus_binds_tighter_than_exponent
# -2**2 is (-2)**2 in bash, unlike most languages.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((-2 ** 2)) -eq 4 ] || fail "got $((-2 ** 2))"
[ $((-(2 ** 2))) -eq -4 ] || fail "got $((-(2 ** 2)))"
echo PASS
exit 0
