#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_multiple_consecutive_underscores
# Multiple consecutive underscores like '____' form a distinct valid variable identifier.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
____="quad_underscore"
[ "$____" = "quad_underscore" ] || fail "multiple consecutive underscores identifier failed"
echo PASS
exit 0
