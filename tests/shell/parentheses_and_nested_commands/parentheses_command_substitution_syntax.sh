#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_command_substitution_syntax
# The $( ... ) construct executes the enclosed command in a subshell and substitutes its output.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
capture=$(printf 'captured_value\n')
[ "$capture" = "captured_value" ] || fail "command substitution: got [$capture]"
echo PASS
exit 0
