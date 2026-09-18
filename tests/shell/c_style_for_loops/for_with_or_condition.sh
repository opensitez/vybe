#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_or_condition
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
for ((i=0; i<2 || i==5; i++)); do
  break
  i=5

done
[ "$i" -eq 0 ] || fail "i=$i"
echo PASS
exit 0
