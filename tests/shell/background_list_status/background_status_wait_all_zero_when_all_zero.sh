#!/usr/bin/env bash
# vybe-test: bash/background_list_status/background_status_wait_all_zero_when_all_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
: &
wait
[ "$?" -eq 0 ] || fail "wait all should succeed"
echo PASS
exit 0
