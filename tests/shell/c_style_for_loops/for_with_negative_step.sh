#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_negative_step
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s=0
for ((i=3;i>0;i--)); do
  s=$((s+i))
done
[ "$s" -eq 6 ] || fail "s=$s"
echo PASS
exit 0
