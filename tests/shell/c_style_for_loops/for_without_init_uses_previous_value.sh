#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_without_init_uses_previous_value
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=5
for (; i>3; i--); do
  :
done
[ "$i" -eq 3 ] || fail "i=$i"
echo PASS
exit 0
