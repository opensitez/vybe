#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=3
out=$(printf "%d-%s" "$num" "value3")
[ "$out" = "$num-value3" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
