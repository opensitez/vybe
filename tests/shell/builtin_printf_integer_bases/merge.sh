#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "14")
out_oct=$(printf "%o" "14")
out_hex=$(printf "%x" "14")
[ "$out_dec" -eq 14 ] || fail "decimal conversion failed"
if (( 14 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 14 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
