#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/post_increment_and_pre_decrement_status
# Status follows the value returned by the update expression, not the assignment action.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
(( x++ )); st=$?
[ "$st" -eq 0 ] || fail "post increment of 1 returns 1 => status 0"
[ "$x" -eq 2 ] || fail "x should become 2"
x=1
(( --x )); st=$?
[ "$st" -eq 0 ] || fail "pre decrement of 1 returns 0 => status 0?" 
[ "$x" -eq 0 ] || fail "x should become 0"
x=2
(( --x )); st=$?
[ "$st" -eq 0 ] || fail "pre decrement of 2 returns 1 => status 0"
[ "$x" -eq 1 ] || fail "x should become 1"
x=0
(( x-- )); st=$?
[ "$st" -eq 1 ] || fail "post decrement of 0 returns 0 => status 1"
[ "$x" -eq -1 ] || fail "x should become -1"
echo PASS
exit 0
