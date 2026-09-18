#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/array_element_incremented_with_assignment
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=(5 10)
((arr[0] += 2)); a=$?
[ "$a" -eq 0 ] || fail "arr[0]+=2 status"
[ "${arr[0]}" -eq 7 ] || fail "arr[0] got ${arr[0]}"
((arr[1]--)); b=$?
[ "$b" -eq 0 ] || fail "arr[1]-- should be status 0"
[ "${arr[1]}" -eq 9 ] || fail "arr[1] got ${arr[1]}"
[ $((arr[0]++)) -eq 7 ] || fail "arr[0]++ returned wrong"
[ "${arr[0]}" -eq 8 ] || fail "arr[0] final"
echo PASS
exit 0
