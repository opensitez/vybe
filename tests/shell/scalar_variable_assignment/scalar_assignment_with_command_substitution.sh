#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_assignment_with_command_substitution
# The right-hand side of a scalar variable assignment can be a command substitution $( ... ).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
computed=$(printf 'generated_%s' "token")
[ "$computed" = "generated_token" ] || fail "assignment with command substitution: got [$computed]"
echo PASS
exit 0
