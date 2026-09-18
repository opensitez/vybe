#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/wait_n_with_no_children_fails
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
wait -n 2>/dev/null
[ $? -ne 0 ] || fail "wait -n without children should fail"
echo PASS
exit 0
