#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_local_variable_scoping
# Variables declared with 'local' inside a function do not leak to the outer calling scope.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
outer_var="outer_initial"
scope_fn() {
    local outer_var="scoped_internal"
    [ "$outer_var" = "scoped_internal" ] || exit 1
}
scope_fn
st=$?
[ "$st" -eq 0 ] || fail "function execution failed"
[ "$outer_var" = "outer_initial" ] || fail "local variable leaked to outer scope: got [$outer_var]"
echo PASS
exit 0
