#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_keyword_without_parentheses
# The Bash-specific 'function name { ... }' syntax defines a function without parentheses.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
function kw_no_parens {
    printf 'kw_no_parens_ok\n'
}
res=$(kw_no_parens)
[ "$res" = "kw_no_parens_ok" ] || fail "function keyword without parens failed: got [$res]"
echo PASS
exit 0
