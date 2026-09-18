#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/escaped
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "7")
out_oct=$(printf "%o" "7")
out_hex=$(printf "%x" "7")
[ "$out_dec" -eq 7 ] || fail "decimal conversion failed"
if (( 7 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 7 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
