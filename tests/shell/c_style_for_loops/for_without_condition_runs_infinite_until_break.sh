#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_without_condition_runs_infinite_until_break
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
for ((; ; )); do
  i=$((i+1))
  [ "$i" -eq 4 ] && break

done
[ "$i" -eq 4 ] || fail "i=$i"
echo PASS
exit 0
