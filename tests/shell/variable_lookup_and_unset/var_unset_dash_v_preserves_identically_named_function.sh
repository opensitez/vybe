#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_dash_v_preserves_identically_named_function
# 'unset -v name' removes the variable without deleting a function that shares the same name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shared_entity="variable_instance"
shared_entity() { printf 'function_instance\n'; }
unset -v shared_entity
[[ ! -v shared_entity ]] || fail "variable was not unset"
res=$(shared_entity)
[ "$res" = "function_instance" ] || fail "function was accidentally destroyed by unset -v: got [$res]"
echo PASS
exit 0
