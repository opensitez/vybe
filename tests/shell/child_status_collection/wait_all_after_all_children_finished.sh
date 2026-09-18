#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_all_after_all_children_finished
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
: &
wait
wait
[ $? -eq 0 ] || fail "second wait after all done should be 0"
echo PASS
exit 0
