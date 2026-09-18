#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_contains_alphanumeric_and_underscores
# Variable identifiers can contain alphanumeric characters and underscores in any subsequent position.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var_1_A_2_b="complex_id"
[ "$var_1_A_2_b" = "complex_id" ] || fail "mixed alphanumeric identifier failed"
echo PASS
exit 0
