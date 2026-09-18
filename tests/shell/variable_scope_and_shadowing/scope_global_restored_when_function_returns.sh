#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_global_restored_when_function_returns
# When a function finishes execution, the shadowed global variable restores its pre-call value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="initial_global"
worker_fn() {
    local x="shadow_temp"
}
worker_fn
[ "$x" = "initial_global" ] || fail "global variable failed to restore: got [$x]"
echo PASS
exit 0
