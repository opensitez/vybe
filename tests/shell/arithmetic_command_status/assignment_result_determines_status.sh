#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/assignment_result_determines_status
# (( x = 0 )) "fails" even though the assignment happened.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( x = 0 )); st=$?
[ "$st" -eq 1 ] || fail "want 1 got $st"
[ "$x" -eq 0 ] || fail "assignment must still happen"
(( x = 3 )); st=$?
[ "$st" -eq 0 ] || fail "want 0 got $st"
echo PASS
exit 0
