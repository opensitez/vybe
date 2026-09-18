#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_c_style_for_loop_header
# Double parentheses in for (( init; test; step )) enclose C-style iteration headers.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
acc=0
for (( k = 1; k <= 4; k++ )); do
    acc=$(( acc + k ))
done
[ "$acc" -eq 10 ] || fail "c-style for loop: want 10, got $acc"
[ "$k" -eq 5 ] || fail "loop variable k after termination: want 5, got $k"
echo PASS
exit 0
