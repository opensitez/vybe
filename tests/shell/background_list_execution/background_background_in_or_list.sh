#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_in_or_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_or_$PPID
: > "$tmp"
false || ( printf b >> "$tmp" ) &
wait
[ "$(cat "$tmp")" = "b" ] || fail "or-list background should execute when lhs fails"
rm -f "$tmp"
echo PASS
exit 0
