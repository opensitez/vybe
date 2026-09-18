#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_ignores_job_control_option_flags
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p=$!
set -m
wait "$p"
set +m
[ $? -eq 0 ] || fail "job control should not affect wait status"
echo PASS
exit 0
