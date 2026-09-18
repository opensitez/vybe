#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/nonzero_result_is_success_zero_is_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 5 )); a=$?
(( 0 )); b=$?
(( -1 )); c=$?
(( 2 > 3 )); d=$?
[ "$a" -eq 0 ] || fail "(( 5 )) want 0 got $a"
[ "$b" -eq 1 ] || fail "(( 0 )) want 1 got $b"
[ "$c" -eq 0 ] || fail "(( -1 )) want 0 got $c"
[ "$d" -eq 1 ] || fail "(( 2 > 3 )) want 1 got $d"
echo PASS
exit 0
