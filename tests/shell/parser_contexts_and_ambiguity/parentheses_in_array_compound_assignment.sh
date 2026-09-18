#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/parentheses_in_array_compound_assignment
# Parentheses in var=(...) are parsed as array compound assignment, not as a subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=initial
x=( one two three )
[ "${#x[@]}" -eq 3 ] || fail "array assignment count: want 3, got ${#x[@]}"
[ "${x[0]}" = "one" ] || fail "array element 0: want 'one', got [${x[0]}]"
echo PASS
exit 0
