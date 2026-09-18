#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_trap_isolation
# Traps configured inside a subshell do not overwrite traps in the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
trap 'printf "parent_exit\n"' EXIT
sub_output=$(
    (
        trap 'printf "child_exit\n"' EXIT
    )
)
[ "$sub_output" = "child_exit" ] || fail "child exit trap did not fire: got [$sub_output]"
trap - EXIT
echo PASS
exit 0
