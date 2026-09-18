#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_with_tilde_expansion
# An unquoted tilde on the RHS of an assignment undergoes tilde expansion to the HOME directory.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
home_path=~
[ "$home_path" = "$HOME" ] || fail "tilde expansion in assignment: want [$HOME], got [$home_path]"
echo PASS
exit 0
