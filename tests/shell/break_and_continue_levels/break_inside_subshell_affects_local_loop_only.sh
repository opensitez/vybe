#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_inside_subshell_affects_local_loop_only
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
result=0
for x in 1 2 3; do
  ( for y in 1 2 3; do
      break
    done )
  result=$((result + 1))
done
[ "$result" -eq 3 ] || fail "parent loop continues result=$result"
echo PASS
exit 0
