#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_exit_status_matches_last_command
# The exit status of a brace group is strictly the exit status of the final command within it.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    (exit 0)
    (exit 12)
}
st1=$?
[ "$st1" -eq 12 ] || fail "brace group status: want 12, got $st1"

{
    (exit 12)
    (exit 0)
}
st2=$?
[ "$st2" -eq 0 ] || fail "brace group status: want 0, got $st2"
echo PASS
exit 0
