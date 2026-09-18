#!/usr/bin/env bash
# vybe-test: bash/parameter_count_and_shift/shift_greater_than_count_fails_and_preserves_list
# Executing 'shift n' where n > $# fails with an error and leaves all parameters unchanged.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "preserved_1" "preserved_2"
shift 10 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "shift greater than \$# should return non-zero exit code"
[ "$#" -eq 2 ] || fail "parameter count altered on failed shift: want 2, got $#"
[ "$1" = "preserved_1" ] && [ "$2" = "preserved_2" ] || fail "parameters altered on failed shift"
echo PASS
exit 0
