#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_nested_subshell_does_not_escape
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
result=0
for x in 1 2 3; do
  ( for y in 1 2 3; do
      continue
    done )
  result=$((result + 1))
done
[ "$result" -eq 3 ] || fail "result=$result"
echo PASS
exit 0
