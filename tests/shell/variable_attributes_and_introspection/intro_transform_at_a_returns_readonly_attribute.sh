#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_a_returns_readonly_attribute
# The ${var@a} transformation contains 'r' for readonly variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly IMMUTABLE="value"
attrs="${IMMUTABLE@a}"
case "$attrs" in
    *r*) : ;;
    *) fail "readonly attribute missing from \${IMMUTABLE@a}: got [$attrs]" ;;
esac
echo PASS
exit 0
