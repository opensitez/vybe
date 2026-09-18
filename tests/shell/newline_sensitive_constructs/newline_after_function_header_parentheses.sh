#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_after_function_header_parentheses
# A newline after 'func()' is permitted before the opening '{' of the function body.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
multiline_fn()
{
    printf 'fn_body_executed\n'
}
res=$(multiline_fn)
[ "$res" = "fn_body_executed" ] || fail "function newline header: got [$res]"
echo PASS
exit 0
