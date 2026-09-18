#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/arithmetic_command_vs_nested_subshells
# (( ... )) is an arithmetic evaluation command, whereas ( ( ... ) ) is nested subshells.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=5
(( x = x + 3 ))
[ "$x" -eq 8 ] || fail "arithmetic evaluation: want 8, got $x"

nested_out=$( ( ( echo "nested" ) ) )
[ "$nested_out" = "nested" ] || fail "nested subshell output: want 'nested', got [$nested_out]"
echo PASS
exit 0
