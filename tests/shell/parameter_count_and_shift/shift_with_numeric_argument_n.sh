#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_with_numeric_argument_n
# The 'shift n' command shifts positional parameters left by n and decrements $# by n.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "a" "b" "c" "d" "e"
shift 3
[ "$#" -eq 2 ] || fail "parameter count after shift 3: want 2, got $#"
[ "$1" = "d" ] || fail "param 1: want 'd', got [$1]"
[ "$2" = "e" ] || fail "param 2: want 'e', got [$2]"
echo PASS
exit 0
