#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_in_while_break_in_for
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
i=0
while [ "$i" -lt 3 ]; do
  for j in 1 2; do
    [ "$j" -eq 2 ] && continue
    count=$((count + 1))
  done
  i=$((i + 1))
  [ "$i" -eq 2 ] && break

done
[ "$count" -eq 3 ] || fail "count $count"
echo PASS
exit 0
