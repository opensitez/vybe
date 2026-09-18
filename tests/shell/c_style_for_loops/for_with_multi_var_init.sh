#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_multi_var_init
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s=0
for ((a=0,b=1; a<3; a++,b+=2)); do
  s=$((s + a + b))
done
[ "$s" -eq 9 ] || fail "s=$s"
echo PASS
exit 0
