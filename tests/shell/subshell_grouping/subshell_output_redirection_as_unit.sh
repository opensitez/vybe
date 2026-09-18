#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_output_redirection_as_unit
# Output redirection attached to a subshell captures stdout from all commands executed in the subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(
    (
        printf 'one\n'
        printf 'two\n'
    )
)
expected=$(printf 'one\ntwo')
[ "$out" = "$expected" ] || fail "subshell redirection output: got [$out]"
echo PASS
exit 0
