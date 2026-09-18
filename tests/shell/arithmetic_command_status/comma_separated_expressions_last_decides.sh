#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/comma_separated_expressions_last_decides
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 0, 1 )); a=$?
(( 1, 0 )); b=$?
[ "$a" -eq 0 ] || fail "(( 0, 1 )) want 0 got $a"
[ "$b" -eq 1 ] || fail "(( 1, 0 )) want 1 got $b"
echo PASS
exit 0
