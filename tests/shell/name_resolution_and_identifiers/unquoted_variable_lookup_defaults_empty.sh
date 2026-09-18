#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/unquoted_variable_lookup_defaults_empty
# Looking up an unset variable identifier without nounset resolves to an empty string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset nonexistent_variable_12345
[ -z "$nonexistent_variable_12345" ] || fail "unset variable should resolve to empty string"
val="${nonexistent_variable_12345}"
[ "$val" = "" ] || fail "explicit unset expansion should be empty string"
echo PASS
exit 0
