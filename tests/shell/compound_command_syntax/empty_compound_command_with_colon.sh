#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/empty_compound_command_with_colon
# A compound command body containing only the null command ':' succeeds with status 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{ :; }
st1=$?
[ "$st1" -eq 0 ] || fail "{ :; } status: want 0, got $st1"

if true; then :; fi
st2=$?
[ "$st2" -eq 0 ] || fail "if-then with : status: want 0, got $st2"

while false; do :; done
st3=$?
[ "$st3" -eq 0 ] || fail "while false with : status: want 0, got $st3"
echo PASS
exit 0
