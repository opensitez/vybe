#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_starts_with_lowercase_letter
# Variable identifiers beginning with lowercase ASCII letters [a-z] are valid.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
valid_lower="alpha"
[ "$valid_lower" = "alpha" ] || fail "lowercase identifier failed"
echo PASS
exit 0
