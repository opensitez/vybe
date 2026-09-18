#!/usr/bin/env bash
# vybe-test: bash/newline_sensitive_constructs/newline_separating_array_literal_elements
# Newlines inside array literal assignment arr=( ... ) separate individual array elements.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=(
    "first element"
    second_element
    third_element
)
[ "${#arr[@]}" -eq 3 ] || fail "array count: want 3, got ${#arr[@]}"
[ "${arr[0]}" = "first element" ] || fail "element 0: want 'first element', got [${arr[0]}]"
[ "${arr[1]}" = "second_element" ] || fail "element 1: want 'second_element', got [${arr[1]}]"
echo PASS
exit 0
