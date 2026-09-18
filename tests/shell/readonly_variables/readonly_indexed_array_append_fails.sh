#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_indexed_array_append_fails
# Appending elements to a readonly array via '+=' fails with an error and non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly -a LIST=( "a" "b" )
( LIST+=( "c" ) ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "array append on readonly should fail"
[ "${#LIST[@]}" -eq 2 ] || fail "array length changed: want 2, got ${#LIST[@]}"
echo PASS
exit 0
