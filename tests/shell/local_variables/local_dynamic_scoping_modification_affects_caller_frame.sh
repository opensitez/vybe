#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_dynamic_scoping_modification_affects_caller_frame
# When a callee assigns to a variable that is local to its caller, it modifies the caller's local copy.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
worker() {
    shared_var="mutated_by_worker"
}
driver() {
    local shared_var="driver_initial"
    worker
    [ "$shared_var" = "mutated_by_worker" ] || fail "caller local was not modified by callee"
}
shared_var="global_safe"
driver
[ "$shared_var" = "global_safe" ] || fail "global variable was accidentally modified: got [$shared_var]"
echo PASS
exit 0
