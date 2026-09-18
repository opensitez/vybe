#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/multiple_background_separators_in_sequence
# Multiple asynchronous & operators can separate successive background jobs on a single line.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(exit 0)& (exit 0)& (exit 0)&
wait
st=$?
[ "$st" -eq 0 ] || fail "wait all background jobs status: want 0, got $st"
echo PASS
exit 0
