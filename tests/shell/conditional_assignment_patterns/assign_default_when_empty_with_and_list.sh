#!/usr/bin/env bash
# vybe-test: bash/conditional_assignment_patterns/assign_default_when_empty_with_and_list
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=
[ -z "$x" ] && x=default
[ "$x" = default ] || fail "empty got default: [$x]"
y=given
[ -z "$y" ] && y=default
[ "$y" = given ] || fail "non-empty kept: [$y]"
echo PASS
exit 0
