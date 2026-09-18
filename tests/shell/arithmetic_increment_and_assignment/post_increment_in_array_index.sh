#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/post_increment_in_array_index
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=(a b c d)
i=1
# index expression with post-increment should use current index then advance
[ "${arr[i++]}" = b ] || fail "first fetch wrong: ${arr[i++]}"
[ "$i" -eq 2 ] || fail "i should be 2"
((i++))
[ "$i" -eq 3 ] || fail "i now 3"
[ "${arr[i]}" = d ] || fail "index after increments wrong"
echo PASS
exit 0
