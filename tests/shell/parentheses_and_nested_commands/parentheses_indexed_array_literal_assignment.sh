#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_indexed_array_literal_assignment
# Parentheses in var=( ... ) define an indexed array literal assignment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "first element" second [4]=fifth )
[ "${#arr[@]}" -eq 3 ] || fail "array length: want 3, got ${#arr[@]}"
[ "${arr[0]}" = "first element" ] || fail "arr[0]: want 'first element', got [${arr[0]}]"
[ "${arr[1]}" = "second" ] || fail "arr[1]: want 'second', got [${arr[1]}]"
[ "${arr[4]}" = "fifth" ] || fail "arr[4]: want 'fifth', got [${arr[4]}]"
echo PASS
exit 0
