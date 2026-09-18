#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_from_failing_background_job
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false &
p=$!
wait "$p"
[ "$?" -eq 1 ] || fail "false should return 1"
echo PASS
exit 0
