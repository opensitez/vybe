#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/basic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "1")
out_oct=$(printf "%o" "1")
out_hex=$(printf "%x" "1")
[ "$out_dec" -eq 1 ] || fail "decimal conversion failed"
if (( 1 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 1 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
