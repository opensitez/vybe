#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_nonexistent_variable_succeeds
# Calling unset on a variable that does not exist succeeds quietly with exit status 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset completely_nonexistent_variable_xyz
st=$?
[ "$st" -eq 0 ] || fail "unsetting nonexistent variable should exit 0, got $st"
echo PASS
exit 0
