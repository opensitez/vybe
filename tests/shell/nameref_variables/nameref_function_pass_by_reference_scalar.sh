#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_function_pass_by_reference_scalar
# A function using 'local -n' receives the caller's variable name and modifies it in-place.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
increment_counter() {
    local -n count_ref=$1
    count_ref=$(( count_ref + 1 ))
}
my_counter=10
increment_counter my_counter
[ "$my_counter" -eq 11 ] || fail "pass-by-reference increment failed: want 11, got $my_counter"
increment_counter my_counter
[ "$my_counter" -eq 12 ] || fail "second pass-by-reference increment failed: want 12, got $my_counter"
echo PASS
exit 0
