#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_brace_group_shares_identical_scope
# Brace groups { ... } execute in the current scope without introducing variable shadowing.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10
{
    x=20
    new_in_brace="created"
}
[ "$x" -eq 20 ] || fail "brace group did not mutate variable in-place: got $x"
[ "$new_in_brace" = "created" ] || fail "variable created inside brace not present in scope"
echo PASS
exit 0
