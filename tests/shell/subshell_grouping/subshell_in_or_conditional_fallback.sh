#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_in_or_conditional_fallback
# A subshell executes as the fallback of '||' and isolates fallback variables from the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="clean"
false || (
    x="fallback_taint"
    exit 0
)
st=$?
[ "$st" -eq 0 ] || fail "subshell fallback status: want 0, got $st"
[ "$x" = "clean" ] || fail "parent x was tainted by subshell fallback: got [$x]"
echo PASS
exit 0
