#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/repeat
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "12")
out_oct=$(printf "%o" "12")
out_hex=$(printf "%x" "12")
[ "$out_dec" -eq 12 ] || fail "decimal conversion failed"
if (( 12 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 12 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
