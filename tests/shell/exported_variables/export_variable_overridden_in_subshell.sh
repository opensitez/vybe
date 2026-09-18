#!/usr/bin/env bash
# vybe-test: bash/exported_variables/export_variable_overridden_in_subshell
# An exported variable overridden in a subshell is passed to child processes spawned from that subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export TOP_EXP="top_level"
sub_res=$(
    (
        export TOP_EXP="subshell_level"
        "$BASH" -c 'printf "%s\n" "$TOP_EXP"'
    )
)
[ "$sub_res" = "subshell_level" ] || fail "subshell override failed: got [$sub_res]"
parent_check=$( "$BASH" -c 'printf "%s\n" "$TOP_EXP"' )
[ "$parent_check" = "top_level" ] || fail "parent export modified by subshell: got [$parent_check]"
echo PASS
exit 0
