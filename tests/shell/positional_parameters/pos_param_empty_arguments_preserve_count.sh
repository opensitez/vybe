#!/usr/bin/env bash
# vybe-test: bash/positional_parameters/pos_param_empty_arguments_preserve_count
# Empty string positional parameters "" occupy valid slots and increment $# accordingly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "start" "" "end"
[ "$#" -eq 3 ] || fail "parameter count: want 3, got $#"
[ "$1" = "start" ] || fail "param 1 mismatch"
[ -z "$2" ] || fail "param 2 should be empty: got [$2]"
[ "$3" = "end" ] || fail "param 3 mismatch"
echo PASS
exit 0
