#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_non_numeric_argument_fails
# Executing shift with a non-numeric argument string fails with a non-zero exit code.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "item1" "item2"
shift "not_a_number" 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "shift with non-numeric argument should fail"
[ "$#" -eq 2 ] || fail "parameter count altered: want 2, got $#"
echo PASS
exit 0
