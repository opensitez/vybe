#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_subshell_inherits_both_exported_and_unexported
# A subshell ( ... ) is a cloned fork that inherits both exported and unexported shell variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
local_scalar="unexported_data"
export global_env="exported_data"
(
    [ "$local_scalar" = "unexported_data" ] || exit 1
    [ "$global_env" = "exported_data" ] || exit 2
)
st=$?
[ "$st" -eq 0 ] || fail "subshell failed to inherit variables: status $st"
echo PASS
exit 0
