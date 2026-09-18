#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/edge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=19
out=$(printf "%d-%s" "$num" "value19")
[ "$out" = "$num-value19" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
