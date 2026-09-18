#!/usr/bin/env bash
# vybe-test: bash/test_builtin_forms/test_does_not_parse_options
# -- is not an end-of-options marker for test; it is just a string, so with
# two arguments it is taken as a unary operator and rejected.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( test -- -f 2>&1 ); st=$?
[ "$st" -eq 2 ] || fail "want status 2 got $st"
[[ $msg == *"unary operator expected"* ]] || fail "got [$msg]"
[ -- = -- ] || fail "-- as a plain operand"
echo PASS
exit 0
