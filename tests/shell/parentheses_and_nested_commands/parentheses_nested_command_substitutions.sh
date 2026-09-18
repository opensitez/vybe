#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_nested_command_substitutions
# Command substitutions can nest cleanly $( ... $( ... ) ... ) without escaping internal parentheses.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(printf 'outer_%s\n' "$(printf 'inner_%s\n' "$(printf 'deep')")")
expected="outer_inner_deep"
[ "$out" = "$expected" ] || fail "nested command substitutions: got [$out]"
echo PASS
exit 0
