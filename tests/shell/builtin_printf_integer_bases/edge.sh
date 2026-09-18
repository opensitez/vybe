#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "19")
out_oct=$(printf "%o" "19")
out_hex=$(printf "%x" "19")
[ "$out_dec" -eq 19 ] || fail "decimal conversion failed"
if (( 19 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 19 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
