#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/consecutive_spaces_collapse_between_args
# Multiple consecutive spaces between words collapse into a single argument separator.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- first     second       third
[ "$#" -eq 3 ] || fail "arg count: want 3, got $#"
[ "$1" = "first" ] || fail "arg 1: want 'first', got [$1]"
[ "$2" = "second" ] || fail "arg 2: want 'second', got [$2]"
[ "$3" = "third" ] || fail "arg 3: want 'third', got [$3]"
echo PASS
exit 0
