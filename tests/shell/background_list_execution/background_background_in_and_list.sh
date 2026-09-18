#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_in_and_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_and_$PPID
: > "$tmp"
: && ( printf a >> "$tmp" ) &
wait
[ "$(cat "$tmp")" = "a" ] || fail "and-list background should execute when lhs succeeds"
rm -f "$tmp"
echo PASS
exit 0
