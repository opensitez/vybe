#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_piped_as_consumer
# A subshell can act as the consumer stage at the end of a pipeline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
printf 'payload_line\n' | (
    read -r line
    [ "$line" = "payload_line" ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell consumer pipeline failed: status $st"
echo PASS
exit 0
