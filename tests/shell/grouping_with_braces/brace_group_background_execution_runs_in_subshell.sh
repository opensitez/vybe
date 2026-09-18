#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_background_execution_runs_in_subshell
# Running a brace group asynchronously with '&' forces execution in a subshell, isolating mutations.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="initial"
{
    var="modified_in_bg"
    exit 0
} &
wait $!
[ "$var" = "initial" ] || fail "asynchronous brace group should not mutate parent: got [$var]"
echo PASS
exit 0
