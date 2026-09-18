#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_c_style_for_loop_variable_leaks_into_scope
# C-style for loop initialization variables leak into the surrounding scope.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset idx
for (( idx = 0; idx < 5; idx++ )); do
    :
done
[ "$idx" -eq 5 ] || fail "c-style for loop index should persist in scope: got $idx"
echo PASS
exit 0
