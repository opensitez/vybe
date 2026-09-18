#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_inherits_parent_functions
# A subshell automatically inherits shell functions defined in the parent environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
parent_greet() {
    printf 'hello_from_parent\n'
}
sub_res=$(
    (
        parent_greet
    )
)
[ "$sub_res" = "hello_from_parent" ] || fail "subshell could not call parent function: got [$sub_res]"
echo PASS
exit 0
