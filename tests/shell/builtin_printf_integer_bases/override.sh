#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/override
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "4")
out_oct=$(printf "%o" "4")
out_hex=$(printf "%x" "4")
[ "$out_dec" -eq 4 ] || fail "decimal conversion failed"
if (( 4 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 4 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
