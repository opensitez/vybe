#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/multiple_ways_of_zeroing_make_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 5-5 )); st=$?
[ "$st" -eq 1 ] || fail "(5-5)=0 should fail"
(( x = 7, x%7 )); st=$?
[ "$st" -eq 1 ] || fail "(x%7) when x=7 should be 0"
(( 6*0 )) ; st=$?
[ "$st" -eq 1 ] || fail "0 product should fail"
echo PASS
exit 0
