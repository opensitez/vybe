#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_inside_while_loop_body
# A brace group inside a while loop body updates condition control variables directly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
k=0
while [ "$k" -lt 3 ]; do
    {
        k=$(( k + 1 ))
    }
done
[ "$k" -eq 3 ] || fail "k inside while-brace: want 3, got $k"
echo PASS
exit 0
