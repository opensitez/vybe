#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/backslash_newline_continuation_removes_both
# A backslash immediately followed by a newline is discarded completely as line continuation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
long_var="first_\
second_\
third"
[ "$long_var" = "first_second_third" ] || fail "backslash newline: want 'first_second_third', got [$long_var]"
echo PASS
exit 0
