#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/decrement_from_negative_starts_at_minus_one
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=-1
[ $((--x)) -eq -2 ] && [ "$x" -eq -2 ] || fail "--x wrong: x=$x"
x=-1
[ $((x--)) -eq -1 ] && [ "$x" -eq -2 ] || fail "x-- wrong: x=$x"
echo PASS
exit 0
