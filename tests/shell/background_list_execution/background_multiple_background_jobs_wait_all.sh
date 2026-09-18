#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_multiple_background_jobs_wait_all
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_multi_$PPID
: > "$tmp"
( printf 1 >> "$tmp" ) &
( printf 2 >> "$tmp" ) &
wait
[ "$(cat "$tmp")" = "12" ] || fail "expected both writes, got [$(cat "$tmp") ]"
rm -f "$tmp"
echo PASS
exit 0
