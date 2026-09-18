#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_and_condition
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
s=0
for ((i=0; i<4 && i!=3; i++)); do
  s=$((s+i))
done
[ "$s" -eq 3 ] || fail "s=$s"
[ "$i" -eq 3 ] || fail "i=$i"
echo PASS
exit 0
