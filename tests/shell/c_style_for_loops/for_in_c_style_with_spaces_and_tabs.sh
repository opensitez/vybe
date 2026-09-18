#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_in_c_style_with_spaces_and_tabs
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s=0
for((	i=0;	i<3;	i++
)); do
  s=$((s+i))
done
[ "$s" -eq 3 ] || fail "s=$s"
echo PASS
exit 0
