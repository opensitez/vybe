#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_indexed_array_element_assignment_fails
# Assigning to an individual element of a readonly indexed array fails with an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly -a ITEMS=( 10 20 30 )
( ITEMS[1]=99 ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "modifying element of readonly array should fail"
[ "${ITEMS[1]}" -eq 20 ] || fail "element was modified: got ${ITEMS[1]}"
echo PASS
exit 0
