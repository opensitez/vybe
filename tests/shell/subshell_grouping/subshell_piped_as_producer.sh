#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_piped_as_producer
# A subshell can act as the producer at the beginning of a pipeline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
piped=$(
    (
        printf 'data_a\n'
        printf 'data_b\n'
    ) | cat
)
expected=$(printf 'data_a\ndata_b')
[ "$piped" = "$expected" ] || fail "subshell producer output: got [$piped]"
echo PASS
exit 0
