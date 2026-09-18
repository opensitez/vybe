#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/arithmetic_ternary_for_min_max
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=7; b=3
max=$(( a > b ? a : b ))
min=$(( a < b ? a : b ))
[ "$max" -eq 7 ] && [ "$min" -eq 3 ] || fail "max=$max min=$min"
(( clamped = a > 5 ? 5 : a ))
[ "$clamped" -eq 5 ] || fail "clamped=$clamped"
echo PASS
exit 0
