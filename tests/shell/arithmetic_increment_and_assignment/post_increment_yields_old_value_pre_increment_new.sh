#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/post_increment_yields_old_value_pre_increment_new
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=5
[ $((x++)) -eq 5 ] && [ "$x" -eq 6 ] || fail "post: x=$x"
[ $((++x)) -eq 7 ] && [ "$x" -eq 7 ] || fail "pre: x=$x"
[ $((x--)) -eq 7 ] && [ "$x" -eq 6 ] || fail "post-dec: x=$x"
[ $((--x)) -eq 5 ] && [ "$x" -eq 5 ] || fail "pre-dec: x=$x"
echo PASS
exit 0
