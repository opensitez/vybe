#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/unset_decrement_starts_at_zero_then_minus_one
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset x
[ $((--x)) -eq 0 ] || fail "--(unset) should return 0"
[ "$x" -eq -1 ] || fail "x should become -1"
unset y
[ $((y--)) -eq 0 ] || fail "y-- should return 0"
[ "$y" -eq -1 ] || fail "y should become -1"
echo PASS
exit 0
