#!/usr/bin/env bash
# vybe-test: bash/variable_existence_tests/dash_v_on_array_name_tests_element_zero
# [[ -v arr ]] is [[ -v arr[0] ]]; use arr[@] to ask whether any element exists.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=(x)
[[ -v a ]] || fail "array with element 0"
b=([1]=y)
[[ -v b ]] && fail "array without element 0 must be false"
[[ -v b[@] ]] || fail "b[@] must be true"
[[ -v b[1] ]] || fail "b[1] must be true"
c=()
[[ -v c[@] ]] && fail "empty array has no elements"
echo PASS
exit 0
