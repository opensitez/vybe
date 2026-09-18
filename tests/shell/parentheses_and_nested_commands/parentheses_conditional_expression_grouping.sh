#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_conditional_expression_grouping
# Inside [[ ... ]], unquoted parentheses ( ... ) group boolean terms with precedence.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[[ ( 1 -eq 1 || 2 -eq 3 ) && 4 -eq 4 ]] || fail "grouped conditional true failed"
[[ ( 1 -eq 2 && 2 -eq 2 ) || 3 -eq 3 ]] || fail "grouped conditional alternation failed"
[[ ( 1 -eq 1 && 2 -eq 3 ) && 4 -eq 4 ]] && fail "grouped conditional false unexpectedly succeeded"
echo PASS
exit 0
