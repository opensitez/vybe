#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_without_increment_uses_manual_update
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
c=0
for ((i=0;i<3;)); do
  c=$((c+1))
  i=$((i+1))
done
[ "$c" -eq 3 ] || fail "c=$c"
