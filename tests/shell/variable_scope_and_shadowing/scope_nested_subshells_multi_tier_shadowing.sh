#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_nested_subshells_multi_tier_shadowing
# Nested subshells each create independent shadow tiers of variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
(
    x=1
    (
        x=2
        [ "$x" -eq 2 ] || exit 1
    )
    [ "$x" -eq 1 ] || exit 2
)
[ "$x" -eq 0 ] || fail "parent x modified: got $x"
echo PASS
exit 0
