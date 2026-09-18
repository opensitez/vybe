#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_from_background_that_writes
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_status_write_$PPID
: > "$tmp"
( printf done >> "$tmp" ) &
p=$!
wait "$p"
[ -f "$tmp" ] || fail "tmp missing"
rm -f "$tmp"
[ "$?" -eq 0 ] || fail "status should be zero"
echo PASS
exit 0
