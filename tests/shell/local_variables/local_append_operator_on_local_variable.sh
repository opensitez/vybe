#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_append_operator_on_local_variable
# The '+=' operator on a local variable appends to the local value without affecting outer variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
outer="initial"
fn_append() {
    local outer="local_part1"
    outer+="_part2"
    [ "$outer" = "local_part1_part2" ] || fail "local append failed: got [$outer]"
}
fn_append
[ "$outer" = "initial" ] || fail "outer variable altered by local +=: got [$outer]"
echo PASS
exit 0
