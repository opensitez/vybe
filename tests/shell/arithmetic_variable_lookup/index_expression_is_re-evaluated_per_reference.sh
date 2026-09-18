#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/index_expression_is_re-evaluated_per_reference
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
vals=(4 8 12)
[ $((vals[i])) -eq 4 ] || fail "vals[0] wrong"
i=1
[ $((vals[i+1])) -eq 12 ] || fail "vals[1+1] wrong"
i=10
[ $((vals[i])) -eq 0 ] || fail "missing index is 0"
ans=$( (( vals[i] )) 2>&1 )
[ "$ans" = "" ] || fail "out-of-range should not print"
echo PASS
exit 0
