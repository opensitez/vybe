#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/test_bracket_requires_surrounding_whitespace
# [ and ] are standard arguments to the test command and require surrounding whitespace.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '[1=1]' 2>/dev/null
st=$?
[ "$st" -eq 127 ] || fail "[1=1] without spaces: want 127, got $st"

[ 1 = 1 ]
st=$?
[ "$st" -eq 0 ] || fail "[ 1 = 1 ] with spaces: want 0, got $st"
echo PASS
exit 0
