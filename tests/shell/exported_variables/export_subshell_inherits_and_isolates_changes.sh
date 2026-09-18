#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_subshell_inherits_and_isolates_changes
# A subshell inherits exported variables and isolating modifications from the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export SYNCED="base_state"
(
    [ "$SYNCED" = "base_state" ] || exit 1
    SYNCED="subshell_mutated"
    [ "$SYNCED" = "subshell_mutated" ] || exit 2
)
st=$?
[ "$st" -eq 0 ] || fail "subshell export inheritance or mutation check failed"
[ "$SYNCED" = "base_state" ] || fail "parent exported variable changed by subshell: got [$SYNCED]"
echo PASS
exit 0
