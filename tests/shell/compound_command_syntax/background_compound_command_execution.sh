#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/background_compound_command_execution
# An entire compound command can be placed in the background with '&'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    printf 'background_step\n' >/dev/null
    exit 0
} &
bg_pid=$!
wait "$bg_pid"
st=$?
[ "$st" -eq 0 ] || fail "background compound command wait status: want 0, got $st"
echo PASS
exit 0
