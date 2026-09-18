#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/arithmetic_compound_command_in_place_mutation
# The (( ... )) compound command mutates shell variables in place using increment and compound operators.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=10
(( x += 5 ))
[ "$x" -eq 15 ] || fail "(( x += 5 )): want 15, got $x"
(( x++ ))
[ "$x" -eq 16 ] || fail "(( x++ )): want 16, got $x"
(( x *= 2 ))
[ "$x" -eq 32 ] || fail "(( x *= 2 )): want 32, got $x"
echo PASS
exit 0
