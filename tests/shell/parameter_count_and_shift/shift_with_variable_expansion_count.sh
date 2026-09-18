#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_with_variable_expansion_count
# The shift command accepts a count supplied via variable expansion '$n'.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "v1" "v2" "v3" "v4"
stride=2
shift "$stride"
[ "$#" -eq 2 ] || fail "count after shift \$stride: want 2, got $#"
[ "$1" = "v3" ] || fail "param 1: want 'v3', got [$1]"
echo PASS
exit 0
