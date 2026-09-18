#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_inside_function_does_not_affect_caller
# Executing shift inside a function shifts the function's arguments without modifying caller's parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "outer_a" "outer_b" "outer_c"
shifting_fn() {
    shift 2
    [ "$#" -eq 0 ] || exit 1
}
shifting_fn "arg1" "arg2"
[ "$#" -eq 3 ] || fail "caller parameter count was modified: want 3, got $#"
[ "$1" = "outer_a" ] && [ "$2" = "outer_b" ] && [ "$3" = "outer_c" ] || fail "caller parameters altered"
echo PASS
exit 0
