#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_break_level_inside_conditional
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
for ((i=0;i<10;i++)); do
  [ "$i" -eq 4 ] && break
  :
done
[ "$i" -eq 4 ] || fail "i=$i"
echo PASS
exit 0
