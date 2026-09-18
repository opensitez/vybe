#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_command_does_not_block_parent_when_still_running
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_nowait_$PPID
: > "$tmp"
( printf one >> "$tmp" ) &
if [ ! -f "$tmp" ]; then fail "tmp should exist immediately"; fi
wait
rm -f "$tmp"
echo PASS
exit 0
