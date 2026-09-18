#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_unset_inside_function_leaves_variable_unset
# Unsetting a local variable inside a function marks it unset within that function's scope.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="outer"
scope_unset() {
    local x="inner"
    unset x
    [[ ! -v x ]] || fail "local x should be unset within function"
    [ -z "$x" ] || fail "local x should expand to empty after unset"
}
scope_unset
[ "$x" = "outer" ] || fail "outer variable altered: got [$x]"
echo PASS
exit 0
