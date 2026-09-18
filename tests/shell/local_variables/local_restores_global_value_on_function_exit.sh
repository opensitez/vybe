#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_restores_global_value_on_function_exit
# When a function finishes execution, shadowed outer variables restore their original values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="original_state"
fn() {
    local var="temporary_override"
}
fn
[ "$var" = "original_state" ] || fail "outer variable not restored after function exit: got [$var]"
echo PASS
exit 0
