#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/stress
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=18
out=$(printf "%d-%s" "$num" "value18")
[ "$out" = "$num-value18" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
