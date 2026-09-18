#!/usr/bin/env bash
# vybe-test: bash/conditional_expression_syntax/cond_parentheses_precedence_grouping
# Unquoted parentheses group subexpressions to override default operator precedence in [[ ... ]].
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ( 1 -eq 2 || 3 -eq 3 ) && 4 -eq 4 ]] || fail "grouped (false || true) && true failed"
[[ 1 -eq 2 && ( 3 -eq 3 || 4 -eq 4 ) ]] && fail "false && (true || true) should be false"
echo PASS
exit 0
