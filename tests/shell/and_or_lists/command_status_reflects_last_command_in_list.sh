#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/command_status_reflects_last_command_in_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: && : && false
s=$?
[ "$s" -eq 1 ] || fail "final status must be from last executed command"
: || : || false
s2=$?
[ "$s2" -eq 1 ] || fail "or chain with two successes and final failure must fail"
echo PASS
exit 0
