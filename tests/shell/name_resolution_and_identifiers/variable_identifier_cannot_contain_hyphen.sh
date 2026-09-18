#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/variable_identifier_cannot_contain_hyphen
# A word containing a hyphen is not a valid variable identifier and cannot be assigned as one.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'my-var=20' 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "my-var=20 should be command not found (127), got $st"
echo PASS
exit 0
