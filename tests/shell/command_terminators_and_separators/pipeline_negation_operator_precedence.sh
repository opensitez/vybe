#!/usr/bin/env bash
# vybe-test: bash/command_terminators_and_separators/pipeline_negation_operator_precedence
# The ! operator negates the exit status of the entire pipeline, not just the first stage.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
! true | false
st=$?
[ "$st" -eq 0 ] || fail "! true | false should exit 0, got $st"

! false | true
st2=$?
[ "$st2" -eq 1 ] || fail "! false | true should exit 1, got $st2"
echo PASS
exit 0
