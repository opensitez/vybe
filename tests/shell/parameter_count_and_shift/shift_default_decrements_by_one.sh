#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_default_decrements_by_one
# The shift command without arguments shifts positional parameters to the left by 1 and decrements $#.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "alpha" "beta" "gamma"
shift
[ "$#" -eq 2 ] || fail "parameter count after shift: want 2, got $#"
[ "$1" = "beta" ] || fail "first param after shift: want 'beta', got [$1]"
[ "$2" = "gamma" ] || fail "second param after shift: want 'gamma', got [$2]"
echo PASS
exit 0
