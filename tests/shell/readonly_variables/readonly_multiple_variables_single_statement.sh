#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_multiple_variables_single_statement
# Multiple variables can be initialized and marked readonly in a single statement.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly R1="v1" R2="v2" R3="v3"
[ "$R1" = "v1" ] && [ "$R2" = "v2" ] && [ "$R3" = "v3" ] || fail "multiple readonly init failed"
( R2="mutated" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "R2 should reject mutation"
echo PASS
exit 0
