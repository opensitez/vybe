#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/comma_operator_evaluates_all_and_returns_last
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
r=$((x=5, y=x*2, y+1))
[ "$r" -eq 11 ] || fail "result want 11 got $r"
[ "$x" -eq 5 ] && [ "$y" -eq 10 ] || fail "assignments inside must persist: x=$x y=$y"
echo PASS
exit 0
