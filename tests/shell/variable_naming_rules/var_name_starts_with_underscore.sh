#!/usr/bin/env bash
# vybe-test: bash/variable_naming_rules/var_name_starts_with_underscore
# Variable identifiers beginning with an underscore character '_' are valid.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
_internal_setting="secure"
[ "$_internal_setting" = "secure" ] || fail "underscore prefix identifier failed"
echo PASS
exit 0
