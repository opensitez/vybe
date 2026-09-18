#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_condition_with_variable_change_inside_update
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
for ((i=0,j=3;i<j;i++,j--)); do
  :
done
[ "$i" -eq 2 ] || fail "i=$i"
[ "$j" -eq 1 ] || fail "j=$j"
echo PASS
exit 0
