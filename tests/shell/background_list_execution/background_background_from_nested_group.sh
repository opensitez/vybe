#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_from_nested_group
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_nested_$PPID
: > "$tmp"
{ ( printf one >> "$tmp" ) & wait; }
[ -s "$tmp" ] || fail "nested group should run"
rm -f "$tmp"
echo PASS
exit 0
