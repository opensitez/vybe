#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/arithmetic_command_truth_status_zero_is_false
# The (( expression )) compound command returns exit status 1 (false) when the expression evaluates to 0.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 0 ))
st=$?
[ "$st" -eq 1 ] || fail "(( 0 )) exit status: want 1 (false), got $st"

(( 5 - 5 ))
st2=$?
[ "$st2" -eq 1 ] || fail "(( 5 - 5 )) exit status: want 1 (false), got $st2"
echo PASS
exit 0
