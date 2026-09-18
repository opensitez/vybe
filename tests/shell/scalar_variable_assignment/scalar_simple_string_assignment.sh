#!/usr/bin/env bash
# vybe-test: bash/scalar_variable_assignment/scalar_simple_string_assignment
# Basic unquoted scalar variable assignment assigns alphanumeric string literal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var=simple_value
[ "$var" = "simple_value" ] || fail "scalar assignment: want 'simple_value', got [$var]"
echo PASS
exit 0
