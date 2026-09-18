#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_compound_increment_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s=0
for ((i=1;i<10;i*=2)); do
  s=$((s+i))
done
[ "$s" -eq 15 ] || fail "s=$s"
echo PASS
exit 0
