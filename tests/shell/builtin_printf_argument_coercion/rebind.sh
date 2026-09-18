#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=16
out=$(printf "%d-%s" "$num" "value16")
[ "$out" = "$num-value16" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
