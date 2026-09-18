#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_append_assignment_to_unset_variable
# Using '+=' on an unset variable sets the variable to the appended value directly.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset fresh_var
fresh_var+="first_content"
[ "$fresh_var" = "first_content" ] || fail "append to unset failed: got [$fresh_var]"
echo PASS
exit 0
