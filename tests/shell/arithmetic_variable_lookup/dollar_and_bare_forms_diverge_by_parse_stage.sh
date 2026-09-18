#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/dollar_and_bare_forms_diverge_by_parse_stage
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
num='1+2'
[ $(( num * 2 )) -eq 6 ] || fail "bare num"
[ $(( $num * 2 )) -eq 5 ] || fail "dollar num"
echo PASS
exit 0
