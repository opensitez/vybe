#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/until_loop_with_continue
# The until loop compound command cycles through iterations using the 'continue' built-in.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
i=0
evens=0
until [ "$i" -ge 6 ]; do
    i=$(( i + 1 ))
    if [ $(( i % 2 )) -ne 0 ]; then
        continue
    fi
    evens=$(( evens + i ))
done
[ "$evens" -eq 12 ] || fail "until continue evens sum: want 12 (2+4+6), got $evens"
echo PASS
exit 0
