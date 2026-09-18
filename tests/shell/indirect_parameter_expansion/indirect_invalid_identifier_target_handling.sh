#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_invalid_identifier_target_handling
# Attempting indirect expansion on a pointer holding an invalid identifier string reports an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ptr="invalid-identifier-name"
( val="${!ptr}" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "indirect expansion on invalid identifier name should return non-zero exit status"
echo PASS
exit 0
