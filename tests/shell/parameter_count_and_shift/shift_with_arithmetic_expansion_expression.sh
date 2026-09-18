#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_with_arithmetic_expansion_expression
# The shift command accepts the result of an inline arithmetic expansion $(( ... )).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- 1 2 3 4 5 6 7
shift $(( 2 + 3 ))
[ "$#" -eq 2 ] || fail "parameter count: want 2, got $#"
[ "$1" -eq 6 ] || fail "param 1: want 6, got $1"
[ "$2" -eq 7 ] || fail "param 2: want 7, got $2"
echo PASS
exit 0
