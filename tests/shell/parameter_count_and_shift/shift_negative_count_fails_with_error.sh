#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_negative_count_fails_with_error
# Executing 'shift -n' fails with an error and non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "a" "b"
shift -1 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "shift with negative count should return non-zero exit code"
[ "$#" -eq 2 ] || fail "parameter count altered on negative shift: want 2, got $#"
echo PASS
exit 0
