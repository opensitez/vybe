#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/modulo_takes_sign_of_dividend
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((7%3)) -eq 1 ] || fail "7%3 got $((7%3))"
[ $((-7%3)) -eq -1 ] || fail "-7%3 want -1 got $((-7%3))"
[ $((7%-3)) -eq 1 ] || fail "7%-3 want 1 got $((7%-3))"
echo PASS
exit 0
