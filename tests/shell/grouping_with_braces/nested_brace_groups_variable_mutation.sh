#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/nested_brace_groups_variable_mutation
# Deeply nested brace groups share the same execution environment and sequentially mutate state.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=0
{
    acc=$(( acc + 1 ))
    {
        acc=$(( acc + 10 ))
        {
            acc=$(( acc + 100 ))
        }
    }
}
[ "$acc" -eq 111 ] || fail "nested brace mutation: want 111, got $acc"
echo PASS
exit 0
