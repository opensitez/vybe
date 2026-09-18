#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_question_mark_updates_on_each_command
# The $? parameter is dynamically recalculated after every executed command statement.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( exit 7 )
st_fail=$?
[ "$st_fail" -eq 7 ] || fail "first status: want 7, got $st_fail"

# A successful command immediately resets $? back to 0
true
[ "$?" -eq 0 ] || fail "\$? was not reset to 0 by subsequent successful command: got $?"
echo PASS
exit 0
