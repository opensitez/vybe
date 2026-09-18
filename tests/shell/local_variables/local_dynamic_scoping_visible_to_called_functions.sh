#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_dynamic_scoping_visible_to_called_functions
# Under Bash dynamic scoping, a caller's local variable is visible to downstream called functions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
child_fn() {
    [ "$dynamic_target" = "caller_context" ] || exit 1
}
parent_fn() {
    local dynamic_target="caller_context"
    child_fn
}
dynamic_target="global_context"
parent_fn
st=$?
[ "$st" -eq 0 ] || fail "dynamic scoping lookup failed in downstream function"
[ "$dynamic_target" = "global_context" ] || fail "global variable was tainted: got [$dynamic_target]"
echo PASS
exit 0
