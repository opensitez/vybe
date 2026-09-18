#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_function_declaration_header
# Empty parentheses '()' after a function name designate a function definition header.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample_fn() {
    printf 'fn_executed\n'
}
res=$(sample_fn)
[ "$res" = "fn_executed" ] || fail "function call: got [$res]"
echo PASS
exit 0
