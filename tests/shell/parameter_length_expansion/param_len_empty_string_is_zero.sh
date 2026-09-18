#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_empty_string_is_zero
# The ${#var} expansion returns 0 when var is set to an empty string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
empty_var=""
[ "${#empty_var}" -eq 0 ] || fail "empty string length: want 0, got ${#empty_var}"
echo PASS
exit 0
