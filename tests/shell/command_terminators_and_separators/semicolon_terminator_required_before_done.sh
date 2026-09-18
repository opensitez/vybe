#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/semicolon_terminator_required_before_done
# When 'done' appears on the same line as the last loop body command, a semicolon terminator is required.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'for x in 1; do echo 1 done' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "loop without terminator before done should fail parsing"

acc=0
for x in 1 2; do acc=$((acc + x)); done
[ "$acc" -eq 3 ] || fail "loop with semicolon before done: want 3, got $acc"
echo PASS
exit 0
