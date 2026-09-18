#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_with_function_body
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_fn_$PPID
emit() { printf fn >> "$tmp"; }
: > "$tmp"
emit &
wait
[ "$(cat "$tmp")" = "fn" ] || fail "function ran in background"
rm -f "$tmp"
echo PASS
exit 0
