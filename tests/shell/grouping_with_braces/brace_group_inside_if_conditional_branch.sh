#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_inside_if_conditional_branch
# A brace group inside an if statement's then or else branch mutates surrounding variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
val=0
if true; then
    {
        step1=1
        step2=2
        val=$(( step1 + step2 ))
    }
fi
[ "$val" -eq 3 ] || fail "val inside if-brace: want 3, got $val"
[ "$step1" -eq 1 ] && [ "$step2" -eq 2 ] || fail "step variables not preserved outside brace group"
echo PASS
exit 0
