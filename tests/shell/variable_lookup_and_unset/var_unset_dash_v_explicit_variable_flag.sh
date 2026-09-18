#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_dash_v_explicit_variable_flag
# The 'unset -v name' explicitly specifies that the target is a variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sample="to_be_removed"
unset -v sample
[[ ! -v sample ]] || fail "unset -v failed to remove variable"
echo PASS
exit 0
