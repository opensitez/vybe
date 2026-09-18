#!/usr/bin/env bash
# vybe-test: bash/c_style_for_loops/for_with_empty_init_uses_existing_var
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=1
for (; i<4; i++); do
  :
done
[ "$i" -eq 4 ] || fail "i=$i"
echo PASS
exit 0
