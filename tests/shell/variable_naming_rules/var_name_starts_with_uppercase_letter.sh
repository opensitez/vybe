#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_starts_with_uppercase_letter
# Variable identifiers beginning with uppercase ASCII letters [A-Z] are valid.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
GLOBAL_CONFIG="active"
[ "$GLOBAL_CONFIG" = "active" ] || fail "uppercase identifier failed"
echo PASS
exit 0
