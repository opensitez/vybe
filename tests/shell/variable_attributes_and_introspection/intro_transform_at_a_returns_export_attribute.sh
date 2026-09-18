#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_a_returns_export_attribute
# The ${var@a} transformation contains 'x' for exported environment variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export EXPORTED_VAR="env_data"
attrs="${EXPORTED_VAR@a}"
case "$attrs" in
    *x*) : ;;
    *) fail "export attribute missing from \${EXPORTED_VAR@a}: got [$attrs]" ;;
esac
echo PASS
exit 0
