#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/command_status_in_arithmetic_pipeline_not_relevant
# Arithmetic command status is independent of stdout/stderr and still determines branching.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 0 )) ; a=$?
printf '%s' '' | :
b=$?
[ "$b" -eq 0 ] || fail "pipeline status must be ignored in this check"
[ "$a" -eq 1 ] || fail "((0)) must fail"
[ "$a" -eq 1 ] || fail "redundant"
echo PASS
exit 0
