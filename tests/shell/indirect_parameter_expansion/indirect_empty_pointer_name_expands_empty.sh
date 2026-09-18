#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_empty_pointer_name_expands_empty
# Indirect expansion when pointer is empty or null produces an error in Bash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ptr=""
( val="${!ptr}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "indirect expansion on empty variable name should fail"
echo PASS
exit 0
