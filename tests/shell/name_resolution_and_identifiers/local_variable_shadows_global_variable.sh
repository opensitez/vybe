#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/local_variable_shadows_global_variable
# A variable declared with 'local' inside a function shadows a global variable of the same identifier.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
score=100
test_scope() {
    local score=50
    [ "$score" -eq 50 ] || fail "local score: want 50, got $score"
}
test_scope
[ "$score" -eq 100 ] || fail "global score: want 100, got $score"
echo PASS
exit 0
