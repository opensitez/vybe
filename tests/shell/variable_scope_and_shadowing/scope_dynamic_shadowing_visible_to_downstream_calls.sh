#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_dynamic_shadowing_visible_to_downstream_calls
# Under dynamic scoping, a downstream called function sees the shadowed variable of its caller.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
callee_fn() {
    [ "$scoped_env" = "caller_shadow" ] || exit 1
}
caller_fn() {
    local scoped_env="caller_shadow"
    callee_fn
}
scoped_env="global_root"
caller_fn
st=$?
[ "$st" -eq 0 ] || fail "callee failed to see caller's shadowed variable"
echo PASS
exit 0
