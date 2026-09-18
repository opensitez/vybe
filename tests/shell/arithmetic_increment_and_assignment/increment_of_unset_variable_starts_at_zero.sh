#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/increment_of_unset_variable_starts_at_zero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset x
[ $((x++)) -eq 0 ] || fail "post-increment of unset yields 0"
[ "$x" -eq 1 ] || fail "x must now be 1, got $x"
unset y
[ $((++y)) -eq 1 ] || fail "pre-increment of unset yields 1"
echo PASS
exit 0
