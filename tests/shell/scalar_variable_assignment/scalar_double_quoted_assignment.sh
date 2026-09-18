#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_double_quoted_assignment
# Double quotes on RHS preserve whitespace while interpolating parameter values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
inner="world"
var="  hello  ${inner}  !  "
[ "$var" = "  hello  world  !  " ] || fail "double quoted assignment failed: got [$var]"
echo PASS
exit 0
