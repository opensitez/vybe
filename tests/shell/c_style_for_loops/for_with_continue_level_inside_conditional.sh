#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_continue_level_inside_conditional
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sum=0
for ((i=0;i<5;i++)); do
  [ "$i" -eq 3 ] && continue
  sum=$((sum+i))
done
[ "$sum" -eq 7 ] || fail "sum=$sum"
echo PASS
exit 0
