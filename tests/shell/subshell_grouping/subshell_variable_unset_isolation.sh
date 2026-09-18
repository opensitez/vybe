#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_variable_unset_isolation
# Unsetting a variable inside a subshell leaves the variable defined and intact in the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
persistent_var="exists"
(
    unset persistent_var
    [ -z "$persistent_var" ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell unset failed"
[ "$persistent_var" = "exists" ] || fail "variable was unset in parent: got [$persistent_var]"
echo PASS
exit 0
