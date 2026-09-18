#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_complex_init_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=1
b=2
for ((i=a+b;i<8;i=i+1)); do
  :
done
[ "$i" -eq 8 ] || fail "i=$i"
echo PASS
exit 0
