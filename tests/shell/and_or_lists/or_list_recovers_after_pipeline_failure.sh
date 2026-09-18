#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/or_list_recovers_after_pipeline_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
(false) || ((count += 1))
[ "$count" -eq 1 ] || fail "rhs should run after subshell failure"
count=0
(:) || ((count += 1))
[ "$count" -eq 0 ] || fail "rhs should not run after subshell success"
echo PASS
exit 0
