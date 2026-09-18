#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_assignment_mutation_fails
# Attempting to assign a new value to a readonly variable produces an error and non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly FIXED="original"
( FIXED="mutated" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "reassigning readonly variable should return non-zero exit code"
[ "$FIXED" = "original" ] || fail "readonly variable was altered: got [$FIXED]"
echo PASS
exit 0
