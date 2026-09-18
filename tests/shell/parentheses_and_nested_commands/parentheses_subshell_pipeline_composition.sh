#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_subshell_pipeline_composition
# Subshells can form both ends of a pipeline simultaneously: ( producer ) | ( consumer ).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=$(
    (
        printf 'data_1\n'
        printf 'data_2\n'
    ) | (
        read -r a
        read -r b
        printf '%s+%s\n' "$a" "$b"
    )
)
[ "$res" = "data_1+data_2" ] || fail "subshell to subshell pipeline: got [$res]"
echo PASS
exit 0
