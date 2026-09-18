#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/bitwise_shift_assignments
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
[ $((x<<=4)) -eq 16 ] || fail "x<<=4"
[ "$x" -eq 16 ] || fail "x after <<=4"
[ $((x>>=2)) -eq 4 ] || fail "x>>=2"
[ "$x" -eq 4 ] || fail "x after >>=2"
[ $((x&=3)) -eq 0 ] || fail "x&=3"
[ "$x" -eq 0 ] || fail "x should be 0"
echo PASS
exit 0
