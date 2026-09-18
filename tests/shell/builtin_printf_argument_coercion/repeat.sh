#!/usr/bin/env bash
# vybe-test: bash/builtin_printf_argument_coercion/repeat
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num=12
out=$(printf "%d-%s" "$num" "value12")
[ "$out" = "$num-value12" ] || fail "printf numeric/string coercion mismatch"
out2=$(printf "%.0f" "$num")
[ "$out2" -eq "$num" ] || fail "float coercion mismatch"
echo PASS
exit 0
