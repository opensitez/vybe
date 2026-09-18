#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/mixed
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=17
out=$(printf "%d-%s" "$num" "value17")
[ "$out" = "$num-value17" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
