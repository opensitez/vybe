#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_subshell_shadows_environment_isolated
# Modifying or assigning a variable inside a subshell creates a shadowed process copy without affecting parent.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="parent_value"
(
    var="subshell_shadow"
    [ "$var" = "subshell_shadow" ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell variable assignment failed"
[ "$var" = "parent_value" ] || fail "parent variable was modified by subshell: got [$var]"
echo PASS
exit 0
