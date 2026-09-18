#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_assignment_via_local_builtin
# The local builtin inside a function evaluates dynamic variable assignments in local scope.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fn_dynamic_local() {
    local key="local_dynamic"
    local "$key=scoped_payload"
    [ "$local_dynamic" = "scoped_payload" ] || exit 1
}
fn_dynamic_local
st=$?
[ "$st" -eq 0 ] || fail "local dynamic assignment failed"
[[ ! -v local_dynamic ]] || fail "dynamic local leaked to outer scope"
echo PASS
exit 0
