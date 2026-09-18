#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_background_in_while_loop_iter
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
tmp=/tmp/bash_bg_exec_while_$PPID
: > "$tmp"
i=0
while [ "$i" -lt 2 ]; do
  ( printf "$i" >> "$tmp" ) &
  i=$((i+1))
done
wait
[ "$(cat "$tmp")" = "01" ] || fail "got $(cat "$tmp")"
rm -f "$tmp"
echo PASS
exit 0
