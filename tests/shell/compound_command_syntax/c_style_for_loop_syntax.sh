#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/c_style_for_loop_syntax
# The C-style for (( init; cond; step )) compound command syntax executes arithmetic iterations.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
total=0
for (( i = 0; i < 5; i++ )); do
    total=$(( total + i ))
done
[ "$total" -eq 10 ] || fail "c-style for total: want 10, got $total"
[ "$i" -eq 5 ] || fail "loop variable i after exit: want 5, got $i"
echo PASS
exit 0
