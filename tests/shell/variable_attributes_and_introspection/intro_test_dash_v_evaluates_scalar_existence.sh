#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_test_dash_v_evaluates_scalar_existence
# The [[ -v var ]] conditional expression checks whether a scalar variable exists (even if empty).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_scalar=""
[[ -v empty_scalar ]] || fail "[[ -v empty_scalar ]] should be true for empty set variable"

unset unassigned_scalar
[[ ! -v unassigned_scalar ]] || fail "[[ -v unassigned_scalar ]] should be false for unset variable"
echo PASS
exit 0
