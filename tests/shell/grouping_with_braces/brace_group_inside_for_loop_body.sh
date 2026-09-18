#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_inside_for_loop_body
# A brace group inside a for loop body executes cleanly across iterations and accumulates values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
sum=0
for n in 1 2 3 4; do
    {
        sum=$(( sum + n ))
    }
done
[ "$sum" -eq 10 ] || fail "sum inside for-brace: want 10, got $sum"
echo PASS
exit 0
