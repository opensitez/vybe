#!/usr/bin/env bash
# vybe-test: bash/brace_expansion/brace_comma_separated_list
# The {a,b,c} brace expansion generates distinct words for each comma-separated member.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- {apple,banana,cherry}
[ "$#" -eq 3 ] || fail "word count: want 3, got $#"
[ "$1" = "apple" ] && [ "$2" = "banana" ] && [ "$3" = "cherry" ] || fail "members mismatch"
echo PASS
exit 0
