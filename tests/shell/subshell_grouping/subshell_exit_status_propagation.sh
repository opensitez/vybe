#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_exit_status_propagation
# The parent shell receives the exact numeric exit status returned by the subshell command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( (exit 0) )
[ "$?" -eq 0 ] || fail "subshell status 0 failed"

( (exit 77) )
[ "$?" -eq 77 ] || fail "subshell status 77 failed"

( (exit 125) )
[ "$?" -eq 125 ] || fail "subshell status 125 failed"
echo PASS
exit 0
