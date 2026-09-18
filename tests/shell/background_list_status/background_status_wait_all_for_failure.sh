#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_wait_all_for_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
false &
wait
[ "$?" -eq 1 ] || fail "wait should reflect failure among children"
echo PASS
exit 0
