#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/child
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "11")
out_oct=$(printf "%o" "11")
out_hex=$(printf "%x" "11")
[ "$out_dec" -eq 11 ] || fail "decimal conversion failed"
if (( 11 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 11 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
