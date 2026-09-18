#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/function
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "2")
out_oct=$(printf "%o" "2")
out_hex=$(printf "%x" "2")
[ "$out_dec" -eq 2 ] || fail "decimal conversion failed"
if (( 2 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 2 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
