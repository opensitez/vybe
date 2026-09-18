#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_multi_digit_parameter_count
# Shifting by multi-digit counts (e.g. shift 10) correctly advances through large argument sets.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- a b c d e f g h i j "eleventh" "twelfth"
shift 10
[ "$#" -eq 2 ] || fail "count after shift 10: want 2, got $#"
[ "$1" = "eleventh" ] || fail "param 1: want 'eleventh', got [$1]"
[ "$2" = "twelfth" ] || fail "param 2: want 'twelfth', got [$2]"
echo PASS
exit 0
