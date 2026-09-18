#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_empty_assignment
# An assignment with no right-hand side sets the variable to an empty string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var="initial"
var=
[ -z "$var" ] || fail "empty assignment failed: got [$var]"
[ "${#var}" -eq 0 ] || fail "length should be 0, got ${#var}"
[[ -v var ]] || fail "empty variable should still be set (-v)"
echo PASS
exit 0
