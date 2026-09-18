#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_zero_leaves_parameters_unchanged
# Executing 'shift 0' leaves positional parameters and $# completely unchanged.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "first" "second"
shift 0
[ "$#" -eq 2 ] || fail "parameter count changed: want 2, got $#"
[ "$1" = "first" ] && [ "$2" = "second" ] || fail "parameters modified by shift 0"
echo PASS
exit 0
