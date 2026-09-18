#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_subshell_in_background
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_subshell_$PPID
: > "$tmp"
( printf outer > "$tmp"; ( printf inner >> "$tmp" ) ) &
wait
[[ "$(cat "$tmp")" = "outerinner" ]] || fail "got [$(cat "$tmp") ]"
rm -f "$tmp"
echo PASS
exit 0
