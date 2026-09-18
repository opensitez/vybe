#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_in_background_wait_status
# A subshell launched with '&' runs in background and its status is gathered by 'wait $!'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
    exit 33
) &
bg_pid=$!
wait "$bg_pid"
st=$?
[ "$st" -eq 33 ] || fail "background subshell wait status: want 33, got $st"
echo PASS
exit 0
