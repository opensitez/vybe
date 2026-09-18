#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_exit_does_not_terminate_parent
# Calling 'exit' inside a subshell terminates only the subshell; the parent shell resumes execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
reached_after=0
(
    exit 10
)
st=$?
reached_after=1
[ "$reached_after" -eq 1 ] || fail "parent failed to continue after subshell exit"
[ "$st" -eq 10 ] || fail "subshell exit status: want 10, got $st"
echo PASS
exit 0
