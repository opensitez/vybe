#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_successive_single_shifts
# Running successive single shift operations sequentially steps through positional parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "step1" "step2" "step3"
[ "$1" = "step1" ] || fail "initial step1 mismatch"
shift
[ "$1" = "step2" ] || fail "after first shift: want step2, got [$1]"
shift
[ "$1" = "step3" ] || fail "after second shift: want step3, got [$1]"
shift
[ "$#" -eq 0 ] || fail "after third shift: want 0 parameters, got $#"
echo PASS
exit 0
