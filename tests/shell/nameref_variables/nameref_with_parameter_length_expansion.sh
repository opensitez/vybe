#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_with_parameter_length_expansion
# Applying ${#ref} on a nameref expands to the string length of the referenced target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source_str="sample_data"
declare -n ref=source_str
[ "${#ref}" -eq 11 ] || fail "nameref length expansion failed: want 11, got ${#ref}"
echo PASS
exit 0
