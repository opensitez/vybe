#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_in_while_loop_multiple_levels
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
i=0
while [ "$i" -lt 2 ]; do
  j=0
  while [ "$j" -lt 3 ]; do
    j=$((j+1))
    [ "$j" -eq 2 ] && continue
    count=$((count+1))
  done
  i=$((i+1))
done
[ "$count" -eq 4 ] || fail "count=$count"
echo PASS
exit 0
