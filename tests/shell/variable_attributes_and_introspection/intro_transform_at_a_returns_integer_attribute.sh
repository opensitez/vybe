#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_a_returns_integer_attribute
# The ${var@a} parameter transformation expands to a string containing the variable's attribute flags.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i count=100
attrs="${count@a}"
case "$attrs" in
    *i*) : ;;
    *) fail "integer attribute missing from \${count@a}: got [$attrs]" ;;
esac
echo PASS
exit 0
