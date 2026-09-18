#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/final
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=20
out=$(printf "%d-%s" "$num" "value20")
[ "$out" = "$num-value20" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
