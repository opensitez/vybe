#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/while_loop_with_break.sh
# The while loop compound command terminates immediately upon encountering the 'break' built-in.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
count=0
while true; do
    count=$(( count + 1 ))
    if [ "$count" -ge 3 ]; then
        break
    fi
done
[ "$count" -eq 3 ] || fail "while break count: want 3, got $count"
echo PASS
exit 0
