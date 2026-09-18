#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_in_while_loop_multiple_levels
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
i=0
while [ "$i" -lt 3 ]; do
  j=0
  while [ "$j" -lt 3 ]; do
    [ "$j" -eq 1 ] && break 2
    count=$((count+1))
    j=$((j+1))
  done
  i=$((i+1))
done
[ "$count" -eq 2 ] || fail "count=$count"
echo PASS
exit 0
