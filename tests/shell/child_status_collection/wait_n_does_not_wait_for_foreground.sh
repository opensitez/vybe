#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_n_does_not_wait_for_foreground
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
: >/tmp/child_status_dummy_$PPID
wait -n
[ $? -eq 0 ] || fail "only children should count"
rm -f /tmp/child_status_dummy_$PPID
echo PASS
exit 0
