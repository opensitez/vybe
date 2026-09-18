#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_inner_stops_one_iteration_of_outer
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
for i in 1 2 3; do
  for j in 1 2 3; do
    ((count++))
    [ "$j" -eq 2 ] && break
  done
done
[ "$count" -eq 6 ] || fail "count got $count"
echo PASS
exit 0
