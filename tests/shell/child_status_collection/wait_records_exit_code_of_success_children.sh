#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_records_exit_code_of_success_children
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
p=$!
wait "$p"
[ $? -eq 0 ] || fail "success child status should be 0"
echo PASS
exit 0
