#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_collects_with_wait_all
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
seen=0
( seen=$((seen+1)) ) &
( seen=$((seen+1)) ) &
wait
[ "$seen" -eq 0 ] || fail "parent variable must not be changed by background"
echo PASS
exit 0
