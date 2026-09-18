#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_when_count_is_zero_fails
# Executing shift when $# is 0 fails with a non-zero exit status without terminating execution.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --
shift 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "shift when \$# is 0 should return non-zero status"
[ "$#" -eq 0 ] || fail "parameter count corrupted: got $#"
echo PASS
exit 0
