#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/six_dash_operators_in_single_bracket
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ 3 -eq 3 ] && [ 3 -ne 4 ] && [ 3 -lt 4 ] && [ 3 -le 3 ] && [ 4 -gt 3 ] && [ 3 -ge 3 ] || fail "true cases"
[ 3 -eq 4 ] && fail "-eq false case"
[ 4 -lt 3 ] && fail "-lt false case"
[ 3 -ge 4 ] && fail "-ge false case"
echo PASS
exit 0
