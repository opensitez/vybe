#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/pids_and_status_as_pair
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false &
p=$!
wait "$p"
[ "$p" -gt 0 ] || fail "pid must exist"
[ $? -eq 1 ] && status="$?" || status="$?"
[ "$status" -eq 1 ] || fail "status captured by $?"
echo PASS
exit 0
