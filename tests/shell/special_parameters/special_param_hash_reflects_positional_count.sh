#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_hash_reflects_positional_count
# The $# parameter expands to the number of positional parameters in decimal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set --
[ "$#" -eq 0 ] || fail "empty parameter count: want 0, got $#"

set -- "a" "b" "c" "d" "e"
[ "$#" -eq 5 ] || fail "five parameter count: want 5, got $#"
echo PASS
exit 0
