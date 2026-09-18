#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_escapes_parentheses_as_arguments
# Escaping '(' and ')' passes them as literal arguments rather than creating a subshell compound command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
capture_args() {
    [ "$#" -eq 3 ] || fail "arg count: want 3, got $#"
    [ "$1" = "(" ] || fail "arg 1: want '(', got [$1]"
    [ "$2" = "inner" ] || fail "arg 2: want 'inner', got [$2]"
    [ "$3" = ")" ] || fail "arg 3: want ')', got [$3]"
}
capture_args \( inner \)
echo PASS
exit 0
