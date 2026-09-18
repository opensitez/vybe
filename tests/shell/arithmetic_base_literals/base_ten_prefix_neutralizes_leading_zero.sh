#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/base_ten_prefix_neutralizes_leading_zero
# The idiom for zero-padded input such as dates: 10#08 is 8, while 08 is an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
m=08
[ $((10#$m)) -eq 8 ] || fail "got $((10#$m))"
[ $((10#09 + 1)) -eq 10 ] || fail "got $((10#09 + 1))"
echo PASS
exit 0
