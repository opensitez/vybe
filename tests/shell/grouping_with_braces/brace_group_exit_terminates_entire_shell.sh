#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_exit_terminates_entire_shell
# An 'exit' statement inside a brace group terminates the whole shell process, not just the group.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
    {
        exit 55
    }
    exit 0
)
st=$?
[ "$st" -eq 55 ] || fail "subshell status with brace exit: want 55, got $st"
echo PASS
exit 0
