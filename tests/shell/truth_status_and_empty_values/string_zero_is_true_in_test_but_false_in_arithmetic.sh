#!/usr/bin/env bash
# vybe-test: bash/truth_status_and_empty_values/string_zero_is_true_in_test_but_false_in_arithmetic
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
[ "$x" ] || fail "[ 0 ] is a non-empty string: true"
[[ $x ]] || fail "[[ 0 ]] is true"
(( x )) && fail "(( 0 )) is false"
y=1
(( y )) || fail "(( 1 )) is true"
echo PASS
exit 0
