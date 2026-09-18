#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_posix_syntax_empty_parentheses
# Standard POSIX function definition syntax uses 'name() { ... }' without the function keyword.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
posix_fn() {
    printf 'posix_ok\n'
}
res=$(posix_fn)
[ "$res" = "posix_ok" ] || fail "posix function call failed: got [$res]"
echo PASS
exit 0
