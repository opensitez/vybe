#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/arithmetic_status_errors_are_recoverable
# Failing (( )) command status should be 1 and execution must continue.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 1/0 )) 2>/dev/null; st=$?
[ "$st" -eq 1 ] || fail "division by zero should be status 1"
after=ok
(( 0 )) 2>/dev/null; st2=$?
[ "$st2" -eq 1 ] || fail "next (( )) should still be evaluated"
(( after=0 ))
[ "$after" -eq 0 ] || fail "arithmetic command continues"
[ "$after" -eq 0 ] || fail "after check"
echo PASS
exit 0
