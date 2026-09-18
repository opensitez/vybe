#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/array_elements_are_referenced_without_dollar
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=(10 20 30); i=1
[ $((arr[1] + 1)) -eq 21 ] || fail "literal index got $((arr[1] + 1))"
[ $((arr[i] + arr[i+1])) -eq 50 ] || fail "index expressions got $((arr[i] + arr[i+1]))"
[ $((arr[5])) -eq 0 ] || fail "missing element is 0"
echo PASS
exit 0
