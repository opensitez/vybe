#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_post_and_pre_updates_in_same_clause
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
for ((; i<3; ++i, ++i)); do
  :
done
[ "$i" -eq 4 ] || fail "i=$i"
echo PASS
exit 0
