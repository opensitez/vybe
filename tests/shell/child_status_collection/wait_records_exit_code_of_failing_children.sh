#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_records_exit_code_of_failing_children
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false &
p=$!
wait "$p"
[ $? -eq 1 ] || fail "failing child status should be preserved"
echo PASS
exit 0
