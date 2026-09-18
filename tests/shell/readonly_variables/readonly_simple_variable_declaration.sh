#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_simple_variable_declaration
# Declaring a variable with the 'readonly' builtin initializes it as an immutable value.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly CONSTANT_VAR="immutable_value"
[ "$CONSTANT_VAR" = "immutable_value" ] || fail "readonly variable assignment failed: got [$CONSTANT_VAR]"
echo PASS
exit 0
