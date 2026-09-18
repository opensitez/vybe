#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/quoting
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "6")
out_oct=$(printf "%o" "6")
out_hex=$(printf "%x" "6")
[ "$out_dec" -eq 6 ] || fail "decimal conversion failed"
if (( 6 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 6 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
