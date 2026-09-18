#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_wait_by_stored_pid
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false &
p=$!
wait "$p"
[ "$?" -eq 1 ] || fail "stored pid should wait to failure"
echo PASS
exit 0
