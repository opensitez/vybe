#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_multilevel_shadowing_three_deep
# Multi-level function calls establish distinct shadowing scopes across each level of invocation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tier="level_0"
level_3() {
    local tier="level_3"
    [ "$tier" = "level_3" ] || exit 1
}
level_2() {
    local tier="level_2"
    level_3
    [ "$tier" = "level_2" ] || exit 2
}
level_1() {
    local tier="level_1"
    level_2
    [ "$tier" = "level_1" ] || exit 3
}
level_1
[ "$tier" = "level_0" ] || fail "level 0 corrupted: got [$tier]"
echo PASS
exit 0
