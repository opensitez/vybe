#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_pipeline_starts_process
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_pipe_$PPID
: > "$tmp"
( printf A >> "$tmp" ) | ( printf B >> "$tmp" ) &
pid=$!
wait "$pid"
[ -f "$tmp" ] || fail "tmp missing"
[ -s "$tmp" ] || fail "pipeline child had no output"
rm -f "$tmp"
echo PASS
exit 0
