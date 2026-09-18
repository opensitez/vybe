#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/variable_identifier_cannot_contain_period
# A word containing a period is not a valid variable identifier and triggers command execution instead.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'my.var=30' 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "my.var=30 should be command not found (127), got $st"
echo PASS
exit 0
