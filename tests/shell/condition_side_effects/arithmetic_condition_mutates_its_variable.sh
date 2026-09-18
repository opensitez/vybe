#!/usr/bin/env bash
# vybe-test: bash/condition_side_effects/arithmetic_condition_mutates_its_variable
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
if (( x++ )); then r=t; else r=f; fi
[ "$r" = f ] || fail "old value 0 is false"
[ "$x" -eq 1 ] || fail "x incremented to $x"
n=0
[ $((n+=2)) -gt 3 ] || :
[ "$n" -eq 2 ] || fail "[ ] operand expansion mutated n=$n"
echo PASS
exit 0
