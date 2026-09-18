#!/usr/bin/env bash
# vybe-test: bash/background_list_status/wait_stores_job_status_without_clobber
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=5
: &
p=$!
wait "$p"
[ "$a" -eq 5 ] || fail "unrelated var unchanged"
[ "$?" -eq 0 ] || fail "status zero"
echo PASS
exit 0
