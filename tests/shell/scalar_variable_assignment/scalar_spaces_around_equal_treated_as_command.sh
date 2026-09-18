#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_spaces_around_equal_treated_as_command
# Including spaces around '=' causes Bash to treat the left operand as a command name, not an assignment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'var = value' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "var = value should fail as command 'var' not found"
echo PASS
exit 0
