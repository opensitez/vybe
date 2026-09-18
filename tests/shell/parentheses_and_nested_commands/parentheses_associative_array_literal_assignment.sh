#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_associative_array_literal_assignment
# Parentheses in declare -A map=( [k]=v ... ) define an associative array literal assignment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A dict=( [name]="Alice" [role]="Admin" [id]=101 )
[ "${#dict[@]}" -eq 3 ] || fail "dict size: want 3, got ${#dict[@]}"
[ "${dict[name]}" = "Alice" ] || fail "dict[name]: want 'Alice', got [${dict[name]}]"
[ "${dict[role]}" = "Admin" ] || fail "dict[role]: want 'Admin', got [${dict[role]}]"
[ "${dict[id]}" = "101" ] || fail "dict[id]: want '101', got [${dict[id]}]"
echo PASS
exit 0
