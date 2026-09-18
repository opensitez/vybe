#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_nested_callee_shifting_does_not_affect_caller
# When a function calls another function that shifts its parameters, the intermediate caller is untouched.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
leaf_fn() {
    shift 3
}
mid_fn() {
    leaf_fn "x" "y" "z"
    [ "$#" -eq 2 ] || fail "mid_fn parameters modified by leaf_fn shift: want 2, got $#"
    [ "$1" = "mid1" ] && [ "$2" = "mid2" ] || fail "mid_fn parameters altered"
}
mid_fn "mid1" "mid2"
echo PASS
exit 0
