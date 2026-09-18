#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_nameref_target_length
# The ${#ref} expansion calculates the character length of the target variable pointed to by ref.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target_string="supercalifragilistic"
declare -n ref=target_string
[ "${#ref}" -eq 20 ] || fail "nameref target length: want 20, got ${#ref}"
echo PASS
exit 0
