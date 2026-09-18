#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/child_status_collection_with_zero_jobs
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
wait 2>/dev/null
[ "$?" -ne 0 ] || fail "no jobs should fail"
echo PASS
exit 0
