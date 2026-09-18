#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/collect_background_command_status_not_parent_exit
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=0
: &
wait
[ "$a" -eq 0 ] || fail "background collection should not alter unrelated parent var"
echo PASS
exit 0
