#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "16")
out_oct=$(printf "%o" "16")
out_hex=$(printf "%x" "16")
[ "$out_dec" -eq 16 ] || fail "decimal conversion failed"
if (( 16 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 16 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
