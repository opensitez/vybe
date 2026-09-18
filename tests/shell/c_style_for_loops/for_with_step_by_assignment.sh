#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_step_by_assignment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
vals=()
for ((i=0;i<10;i=i+3)); do
  vals+=("$i")
done
[ "${#vals[@]}" -eq 4 ] || fail "len=${#vals[@]}"
[ "${vals[3]}" = 9 ] || fail "last=${vals[3]}"
echo PASS
exit 0
