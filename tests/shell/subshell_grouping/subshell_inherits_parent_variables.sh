#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_inherits_parent_variables
# A subshell inherits copies of both exported and unexported shell variables from its parent.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
parent_unexported="local_data"
export parent_exported="env_data"
(
    [ "$parent_unexported" = "local_data" ] || exit 1
    [ "$parent_exported" = "env_data" ] || exit 2
)
st=$?
[ "$st" -eq 0 ] || fail "subshell variable inheritance failed: status $st"
echo PASS
exit 0
