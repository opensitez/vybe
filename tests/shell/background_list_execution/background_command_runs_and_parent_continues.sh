#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_command_runs_and_parent_continues
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_parent_$PPID
: > "$tmp"
( printf ok >> "$tmp" ) &
wait
[[ -f "$tmp" ]] || fail "tmp missing"
[[ -s "$tmp" ]] || fail "background command didn't write"
rm -f "$tmp"
echo PASS
exit 0
