#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_quoted_inside_strings
# Parentheses enclosed within double or single quotes are literal characters without syntactic meaning.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s1="(hello)"
s2='(world)'
[ "$s1" = "(hello)" ] || fail "double quoted parens: got [$s1]"
[ "$s2" = "(world)" ] || fail "single quoted parens: got [$s2]"
echo PASS
exit 0
