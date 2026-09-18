#!/usr/bin/env bash
# vybe-test: bash/parser_contexts_and_ambiguity/arithmetic_command_truth_status_nonzero_is_true
# The (( expression )) compound command returns exit status 0 (true) when the expression evaluates non-zero.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(( 1 ))
st=$?
[ "$st" -eq 0 ] || fail "(( 1 )) exit status: want 0 (true), got $st"

(( -42 ))
st2=$?
[ "$st2" -eq 0 ] || fail "(( -42 )) exit status: want 0 (true), got $st2"
echo PASS
exit 0
