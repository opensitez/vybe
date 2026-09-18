#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_arithmetic_comparison
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
max=2
sum=0
for ((i=0; i<=max; i++)); do
  sum=$((sum+i))
done
[ "$sum" -eq 3 ] || fail "sum=$sum"
echo PASS
exit 0
