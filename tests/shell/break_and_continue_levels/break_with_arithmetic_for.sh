#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_with_arithmetic_for
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s=0
for ((i=0;i<5;i++)); do
  [ "$i" -eq 3 ] && break
  s=$((s+1))
 done
[ "$s" -eq 3 ] || fail "s=$s"
echo PASS
exit 0
