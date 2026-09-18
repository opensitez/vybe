#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_unset_variable_is_zero
# The ${#var} expansion returns 0 when var is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset missing_var
[ "${#missing_var}" -eq 0 ] || fail "unset variable length: want 0, got ${#missing_var}"
echo PASS
exit 0
