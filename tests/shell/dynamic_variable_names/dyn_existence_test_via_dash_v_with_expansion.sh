#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_existence_test_via_dash_v_with_expansion
# The test [[ -v $pointer ]] dynamically checks whether the variable named by $pointer exists.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
actual_slot="exists"
p_set="actual_slot"
p_unset="nonexistent_slot_xyz"
[[ -v $p_set ]] || fail "dynamic -v test failed for existing variable"
[[ -v $p_unset ]] && fail "dynamic -v test should return false for nonexistent variable"
echo PASS
exit 0
