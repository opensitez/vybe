#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_arithmetic_mutation_fails
# Arithmetic operations that mutate variables like '(( ro++ ))' fail on readonly variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly NUM=10
( (( NUM++ )) ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "(( NUM++ )) on readonly should return non-zero exit status"
[ "$NUM" -eq 10 ] || fail "readonly number altered: got $NUM"
echo PASS
exit 0
