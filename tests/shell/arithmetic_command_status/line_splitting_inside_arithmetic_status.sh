#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/line_splitting_inside_arithmetic_status
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=0
(( 1 + \
   2 + \
   3 )); st=$?
[ "$st" -eq 0 ] || fail "multiline arithmetic should succeed"
[ "$x" -eq 0 ] || fail "no side effect expected in pure expression"
(( \
  2 * \
  0 )); st=$?
[ "$st" -eq 1 ] || fail "multiline zero expression should fail"
echo PASS
exit 0
