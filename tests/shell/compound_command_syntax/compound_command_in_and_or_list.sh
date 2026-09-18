#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/compound_command_in_and_or_list
# Compound commands participate seamlessly in && and || conditional lists.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
false || {
    x=1
    y=2
}
[ "$x" -eq 1 ] && [ "$y" -eq 2 ] || fail "|| compound command failed"

z=0
true && {
    z=10
}
[ "$z" -eq 10 ] || fail "&& compound command failed"
echo PASS
exit 0
