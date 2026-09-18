#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_indexed_array_mutation_fails
# An indexed array marked readonly rejects reassignment of the entire array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly -a RO_ARR=( "alpha" "beta" "gamma" )
[ "${#RO_ARR[@]}" -eq 3 ] || fail "array init failed"
( RO_ARR=( "new" "elements" ) ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "readonly array reassignment should return non-zero exit status"
echo PASS
exit 0
