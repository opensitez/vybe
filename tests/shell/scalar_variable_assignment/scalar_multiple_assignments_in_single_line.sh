#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_multiple_assignments_in_single_line
# Multiple variable assignments can appear on a single command line separated by spaces.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=1 b=2 c=3
[ "$a" -eq 1 ] || fail "a: want 1, got $a"
[ "$b" -eq 2 ] || fail "b: want 2, got $b"
[ "$c" -eq 3 ] || fail "c: want 3, got $c"
echo PASS
exit 0
