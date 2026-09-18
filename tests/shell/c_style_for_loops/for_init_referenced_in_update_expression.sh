#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_init_referenced_in_update_expression
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
base=1
for ((i=base; i<4; i=i+base)); do
  :
done
[ "$i" -eq 4 ] || fail "i=$i"
