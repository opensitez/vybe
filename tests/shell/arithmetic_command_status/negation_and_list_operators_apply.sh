#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/negation_and_list_operators_apply
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
! (( 0 )); st=$?
[ "$st" -eq 0 ] || fail "! (( 0 )) want 0 got $st"
x=2
(( x > 1 )) && r=yes || r=no
[ "$r" = yes ] || fail "&& after (( )): got [$r]"
(( x > 5 )) || r=fallback
[ "$r" = fallback ] || fail "|| after (( )): got [$r]"
echo PASS
exit 0
