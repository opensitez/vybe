#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_variable_mutation_isolation
# Mutating an existing variable inside a subshell ( ... ) leaves the parent shell variable unchanged.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="initial"
(
    x="mutated"
    [ "$x" = "mutated" ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell failed: status $st"
[ "$x" = "initial" ] || fail "parent x was modified: got [$x]"
echo PASS
exit 0
