#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_question_mark_exit_status_success_and_failure
# The $? parameter expands to the decimal exit status of the most recently executed foreground pipeline.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( exit 0 )
[ "$?" -eq 0 ] || fail "\$? after exit 0: want 0, got $?"

( exit 42 )
st=$?
[ "$st" -eq 42 ] || fail "\$? after exit 42: want 42, got $st"
echo PASS
exit 0
