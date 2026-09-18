#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/merge
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=14
out=$(printf "%d-%s" "$num" "value14")
[ "$out" = "$num-value14" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
