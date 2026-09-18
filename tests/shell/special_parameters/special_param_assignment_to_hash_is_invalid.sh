#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_assignment_to_hash_is_invalid
# Direct assignment to the special parameter $# via printf -v or export fails with an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
printf -v '#' "10" 2>/dev/null
st1=$?
[ "$st1" -ne 0 ] || fail "printf -v targeting special parameter # should fail"

export '#=10' 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "export targeting special parameter # should fail"
echo PASS
exit 0
