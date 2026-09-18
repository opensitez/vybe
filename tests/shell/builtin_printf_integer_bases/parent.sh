#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/parent
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "10")
out_oct=$(printf "%o" "10")
out_hex=$(printf "%x" "10")
[ "$out_dec" -eq 10 ] || fail "decimal conversion failed"
if (( 10 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 10 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
