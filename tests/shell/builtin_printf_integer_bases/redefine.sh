#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/redefine
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "9")
out_oct=$(printf "%o" "9")
out_hex=$(printf "%x" "9")
[ "$out_dec" -eq 9 ] || fail "decimal conversion failed"
if (( 9 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 9 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
