#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_keyword_with_parentheses
# The hybrid 'function name() { ... }' syntax is valid Bash combining both keyword and parentheses.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
function kw_with_parens() {
    printf 'kw_with_parens_ok\n'
}
res=$(kw_with_parens)
[ "$res" = "kw_with_parens_ok" ] || fail "function keyword with parens failed: got [$res]"
echo PASS
exit 0
