#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_command_substitution_in_arithmetic
# A command substitution $( ... ) nested inside arithmetic evaluation $(( ... )) is evaluated first.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
total=$(( 10 + $(printf '25\n') * 2 ))
[ "$total" -eq 60 ] || fail "cmdsub in arithmetic: want 60 (10 + 50), got $total"
echo PASS
exit 0
