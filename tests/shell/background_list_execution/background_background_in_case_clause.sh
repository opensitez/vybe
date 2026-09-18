#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_in_case_clause
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
mode=branch
tmp=/tmp/bash_bg_exec_case_$PPID
: > "$tmp"
case "$mode" in
  branch)
    ( printf casehit >> "$tmp" ) &
    ;;
esac
wait
[ "$(cat "$tmp")" = "casehit" ] || fail "case branch background should execute"
rm -f "$tmp"
echo PASS
exit 0
