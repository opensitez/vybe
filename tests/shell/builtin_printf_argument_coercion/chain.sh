#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=5
out=$(printf "%d-%s" "$num" "value5")
[ "$out" = "$num-value5" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
