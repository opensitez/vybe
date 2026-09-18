#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/exponent_is_right_associative
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ $((2**3**2)) -eq 512 ] || fail "2**3**2 want 512 got $((2**3**2))"
[ $(((2**3)**2)) -eq 64 ] || fail "(2**3)**2 want 64 got $(((2**3)**2))"
[ $((5**0)) -eq 1 ] || fail "5**0 want 1"
[ $(((-2)**3)) -eq -8 ] || fail "(-2)**3 want -8 got $(((-2)**3))"
echo PASS
exit 0
