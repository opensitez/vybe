#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_explicit_numeric_return
# The 'return N' builtin exits the function immediately with the specified integer status code.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ret_fn() {
    return "$1"
    printf 'unreachable\n'
}
ret_fn 0
[ "$?" -eq 0 ] || fail "return 0 failed"
ret_fn 42
[ "$?" -eq 42 ] || fail "return 42 failed"
ret_fn 127
[ "$?" -eq 127 ] || fail "return 127 failed"
echo PASS
exit 0
