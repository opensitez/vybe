#!/usr/bin/env bash
# vybe-test: bash/test_builtin_forms/return_status_zero_one_two
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ a = a ]; a=$?
[ a = b ]; b=$?
[ a -eq b ] 2>/dev/null; c=$?
[ "$a" -eq 0 ] || fail "true want 0 got $a"
[ "$b" -eq 1 ] || fail "false want 1 got $b"
[ "$c" -eq 2 ] || fail "error want 2 got $c"
echo PASS
exit 0
