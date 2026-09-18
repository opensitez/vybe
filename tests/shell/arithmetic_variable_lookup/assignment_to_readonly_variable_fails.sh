#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/assignment_to_readonly_variable_fails
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly r=1
(( r = 2 )) 2>/dev/null; st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[ "$r" -eq 1 ] || fail "r must be unchanged"
echo PASS
exit 0
