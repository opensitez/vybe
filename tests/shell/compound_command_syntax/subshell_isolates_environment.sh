#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/subshell_isolates_environment
# A subshell compound command ( ... ) executes in an isolated environment and does not mutate parent variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="initial"
(
    x="mutated"
    [ "$x" = "mutated" ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell failed with status $st"
[ "$x" = "initial" ] || fail "parent environment was mutated by subshell: got [$x]"
echo PASS
exit 0
