#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_output_redirection_aggregates_all_stdout
# Applying output redirection to a brace group captures standard output of all internal commands.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(
    {
        printf 'first\n'
        printf 'second\n'
    }
)
expected=$(printf 'first\nsecond')
[ "$out" = "$expected" ] || fail "aggregated stdout: got [$out]"
echo PASS
exit 0
