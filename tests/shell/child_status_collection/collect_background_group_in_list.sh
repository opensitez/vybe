#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/collect_background_group_in_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( : ) &
( : ) &
wait -n
wait -n
[ $? -eq 0 ] || fail "two wait -n should succeed"
echo PASS
exit 0
