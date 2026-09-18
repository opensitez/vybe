#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/assignment_result_with_commas_still_uses_last_value
# Comma sequence status is based only on the rightmost sub-expression result.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
(( x = 1, 0 )); st=$?
[ "$st" -eq 1 ] || fail "want 1 got $st"
[ "$x" -eq 1 ] || fail "assignment must happen, x=$x"
(( x = 0, x + 5 )); st=$?
[ "$st" -eq 0 ] || fail "last expression 5 should be success, got $st"
[ "$x" -eq 0 ] || fail "x must be 0, got $x"
echo PASS
exit 0
