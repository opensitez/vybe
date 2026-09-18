#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/division_truncates_toward_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((7/2)) -eq 3 ] || fail "7/2 want 3 got $((7/2))"
[ $((-7/2)) -eq -3 ] || fail "-7/2 want -3 got $((-7/2))"
[ $((7/-2)) -eq -3 ] || fail "7/-2 want -3 got $((7/-2))"
echo PASS
exit 0
