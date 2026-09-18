#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/semicolon_terminator_required_before_closing_brace
# In a brace group { list; }, the command immediately preceding '}' must end with a semicolon or newline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '{ echo 1 }' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "brace group without terminator should fail parsing"

val=0
{ val=1; }
[ "$val" -eq 1 ] || fail "brace group with semicolon: want 1, got $val"
echo PASS
exit 0
