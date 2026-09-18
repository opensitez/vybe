#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_list_in_grouping
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_group_$PPID
: > "$tmp"
{ printf a >> "$tmp"; printf b >> "$tmp"; } &
wait
[ "$(cat "$tmp")" = "ab" ] || fail "got [$(cat "$tmp") ]"
rm -f "$tmp"
echo PASS
exit 0
