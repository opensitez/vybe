#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_dynamic_unshadowing_when_caller_returns
# When a shadowing caller returns, subsequent calls to the same callee see the restored global scope.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inspect_var() {
    printf '%s\n' "$shared_target"
}
wrapper() {
    local shared_target="shadowed"
    inspect_var
}
shared_target="global_state"
first_call=$(wrapper)
second_call=$(inspect_var)
[ "$first_call" = "shadowed" ] || fail "first call should see shadow: got [$first_call]"
[ "$second_call" = "global_state" ] || fail "second call should see global: got [$second_call]"
echo PASS
exit 0
