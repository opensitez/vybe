#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_integer_attribute_arithmetic_assignment
# The 'declare -i' flag evaluates all subsequent assignments to the variable as arithmetic expressions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i num
num="10 * 3 + 5"
[ "$num" -eq 35 ] || fail "arithmetic evaluation on assignment failed: want 35, got $num"
echo PASS
exit 0
