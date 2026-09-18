#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_from_single_background_job
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p=$!
wait "$p"
[ "$?" -eq 0 ] || fail ": should return 0"
echo PASS
exit 0
