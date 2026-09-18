#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/standard_for_with_init_condition_increment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s=0
for ((i=0;i<3;i++)); do
  s=$((s+i))
done
[ "$s" -eq 3 ] || fail "sum got $s"
echo PASS
exit 0
