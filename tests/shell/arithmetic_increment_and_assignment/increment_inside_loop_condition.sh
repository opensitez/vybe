#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/increment_inside_loop_condition
# The condition is evaluated one extra time, so i ends one past the bound.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0; n=0
while (( i++ < 3 )); do n=$((n+1)); done
[ "$n" -eq 3 ] || fail "body ran $n times"
[ "$i" -eq 4 ] || fail "i want 4 got $i"
echo PASS
exit 0
