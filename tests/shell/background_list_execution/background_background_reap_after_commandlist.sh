#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_reap_after_commandlist
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_reap_$PPID
: > "$tmp"
( printf one >> "$tmp"; printf two >> "$tmp" ) &
wait
grep_marker=$(cat "$tmp")
[ "$grep_marker" = "onetwo" ] || fail "both parts should run"
rm -f "$tmp"
echo PASS
exit 0
