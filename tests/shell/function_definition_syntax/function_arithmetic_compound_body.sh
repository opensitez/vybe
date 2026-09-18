#!/usr/bin/env bash
# vybe-test: bash/function_definition_syntax/function_arithmetic_compound_body
# A function body can be an arithmetic command (( ... )), returning 0 for non-zero and 1 for zero.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
is_positive() (( $1 > 0 ))
is_positive 5
st_pos=$?
[ "$st_pos" -eq 0 ] || fail "is_positive 5: want status 0, got $st_pos"

is_positive -3
st_neg=$?
[ "$st_neg" -eq 1 ] || fail "is_positive -3: want status 1, got $st_neg"
echo PASS
exit 0
