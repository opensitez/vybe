#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_using_pre_increment_update
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
for ((; i<4; ++i)); do
  :
done
[ "$i" -eq 4 ] || fail "i=$i"
echo PASS
exit 0
