#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_conditional_on_unevaluated_init
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=9
for ((; i>5; )); do
  i=$((i-2))
done
[ "$i" -eq 5 ] || fail "i=$i"
echo PASS
exit 0
