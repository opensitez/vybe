#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_with_lastpipe_preserves_variables
# With lastpipe enabled and job control inactive, the last pipeline command runs in the current shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set +m
shopt -s lastpipe
received="unset"
printf 'pipe_payload\n' | { read -r received; }
[ "$received" = "pipe_payload" ] || fail "lastpipe failed to mutate variable: got [$received]"
echo PASS
exit 0
