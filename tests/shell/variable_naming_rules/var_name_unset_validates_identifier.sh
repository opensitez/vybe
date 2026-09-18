#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_unset_validates_identifier
# Calling 'unset -v' with an invalid variable identifier (containing invalid punctuation) reports an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset -v "bad-var-name" 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unset -v bad-var-name should return non-zero exit status"
echo PASS
exit 0
