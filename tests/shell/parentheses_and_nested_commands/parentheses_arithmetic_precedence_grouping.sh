#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_arithmetic_precedence_grouping
# Parentheses inside arithmetic evaluation $(( ... )) override default operator precedence.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
without_parens=$(( 2 + 3 * 4 ))
with_parens=$(( (2 + 3) * 4 ))
[ "$without_parens" -eq 14 ] || fail "without parens: want 14, got $without_parens"
[ "$with_parens" -eq 20 ] || fail "with parens: want 20, got $with_parens"
echo PASS
exit 0
