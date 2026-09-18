#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_modifies_current_shell_variables
# Commands inside a brace group run in the current execution context and mutate outer variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10
y=20
{
    x=50
    y=100
}
[ "$x" -eq 50 ] || fail "x after brace group: want 50, got $x"
[ "$y" -eq 100 ] || fail "y after brace group: want 100, got $y"
echo PASS
exit 0
