#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/compound_command_status_is_last_command
# The exit status of a compound command is the exit status of the last command executed within it.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
{
    (exit 0)
    (exit 7)
}
st=$?
[ "$st" -eq 7 ] || fail "compound status: want 7, got $st"
echo PASS
exit 0
