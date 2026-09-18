#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/nested_c_style_loops_interaction
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
total=0
for ((i=0;i<2;i++)); do
  for ((j=0;j<2;j++)); do
    total=$((total+i+j))
  done

done
[ "$total" -eq 4 ] || fail "total=$total"
echo PASS
exit 0
