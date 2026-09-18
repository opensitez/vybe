#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/variable_identifier_cannot_start_with_digit
# An assignment to a word starting with a digit is not an assignment and attempts command execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '1var=10' 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "1var=10 should be command not found (127), got $st"
echo PASS
exit 0
