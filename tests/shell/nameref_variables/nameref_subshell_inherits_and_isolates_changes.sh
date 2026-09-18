#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_subshell_inherits_and_isolates_changes
# A subshell inherits a nameref; mutations through the nameref inside the subshell are isolated.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
master_data="unchanged"
declare -n ref=master_data
(
    [ "$ref" = "unchanged" ] || exit 1
    ref="subshell_mutated"
    [ "$master_data" = "subshell_mutated" ] || exit 2
)
st=$?
[ "$st" -eq 0 ] || fail "subshell nameref operation failed"
[ "$master_data" = "unchanged" ] || fail "parent target modified by subshell: got [$master_data]"
echo PASS
exit 0
