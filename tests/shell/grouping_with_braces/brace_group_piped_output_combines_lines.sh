#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_piped_output_combines_lines
# Piping out of a brace group combines output from all enclosed commands into a single stream.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
piped=$(
    {
        printf 'item1\n'
        printf 'item2\n'
    } | cat
)
expected=$(printf 'item1\nitem2')
[ "$piped" = "$expected" ] || fail "piped brace group output: got [$piped]"
echo PASS
exit 0
