#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_escaped_as_literal_arguments
# Escaping parentheses \( and \) passes them as literal arguments to commands rather than altering syntax.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- \( arg \)
[ "$#" -eq 3 ] || fail "arg count: want 3, got $#"
[ "$1" = "(" ] || fail "arg 1: want '(', got [$1]"
[ "$2" = "arg" ] || fail "arg 2: want 'arg', got [$2]"
[ "$3" = ")" ] || fail "arg 3: want ')', got [$3]"
echo PASS
exit 0
