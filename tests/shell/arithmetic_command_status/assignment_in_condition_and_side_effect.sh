#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/assignment_in_condition_and_side_effect
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
i=0
while (( (x = x + 1) && x < 3 )); do i=$((i+1)); done
[ "$i" -eq 2 ] || fail "loop body executed wrong times i=$i"
[ "$x" -eq 3 ] || fail "x should finish at 3 got $x"
echo PASS
exit 0
