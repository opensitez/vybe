#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/multiple_consecutive_newlines_collapsed
# Multiple sequential newlines act as a single command separator between statements.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=1




b=2
[ "$((a + b))" -eq 3 ] || fail "multiple newlines: want 3, got $((a + b))"
echo PASS
exit 0
