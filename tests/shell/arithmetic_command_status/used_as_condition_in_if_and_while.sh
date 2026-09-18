#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/used_as_condition_in_if_and_while
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
n=5
if (( n > 3 )); then r=big; else r=small; fi
[ "$r" = big ] || fail "if: got [$r]"
i=0; sum=0
while (( i < 4 )); do (( sum += i, i++ )); done
[ "$sum" -eq 6 ] && [ "$i" -eq 4 ] || fail "while: sum=$sum i=$i"
echo PASS
exit 0
