#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_with_level_two_jumps_outer
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
for i in 1 2 3; do
  for j in 1 2 3; do
    for k in 1 2 3; do
      ((count++))
      break 2
    done
  done
done
[ "$count" -eq 1 ] || fail "count got $count"
echo PASS
exit 0
