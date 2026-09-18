#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_positional_parameter_fails
# Executing 'unset 1' does not affect positional parameter $1; the argument remains intact.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "arg1" "arg2"
unset 1
[ "$1" = "arg1" ] || fail "positional parameter 1 modified by unset 1: got [$1]"
[ "$2" = "arg2" ] || fail "positional parameter 2 modified by unset 1: got [$2]"
[ "$#" -eq 2 ] || fail "parameter count modified by unset 1: want 2, got $#"
echo PASS
exit 0
