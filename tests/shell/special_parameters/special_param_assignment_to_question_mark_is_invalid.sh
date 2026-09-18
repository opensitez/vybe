#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_assignment_to_question_mark_is_invalid
# Direct assignment to the special parameter $? via '?=val' is a syntax error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '?="0"' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "assignment to special parameter ? should fail"
echo PASS
exit 0
