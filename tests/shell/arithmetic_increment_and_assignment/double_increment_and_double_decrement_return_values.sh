#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/double_increment_and_double_decrement_return_values
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=2
[ $((x+=1)) -eq 3 ] || fail "x+=1 got $((x+=1))"
[ "$x" -eq 3 ] || fail "x should be 3"
x=2
out=$((x++, x++, x))
[ "$out" -eq 4 ] || fail "post chain final value got $out"
[ "$x" -eq 4 ] || fail "x should be 4"
x=5
res=$(( --x, --x ))
[ "$res" -eq 3 ] || fail "two pre decrements final got $res"
[ "$x" -eq 3 ] || fail "x should be 3"
echo PASS
exit 0
