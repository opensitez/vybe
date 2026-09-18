#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_integer_bases/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out_dec=$(printf "%d" "3")
out_oct=$(printf "%o" "3")
out_hex=$(printf "%x" "3")
[ "$out_dec" -eq 3 ] || fail "decimal conversion failed"
if (( 3 >= 8 )); then
  [ "$out_oct" != "$out_dec" ] || fail "octal should differ from decimal"
fi
if (( 3 >= 10 )); then
  [ "$out_hex" != "$out_dec" ] || fail "hex should differ from decimal"
fi
echo PASS
exit 0
