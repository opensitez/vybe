#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/valid_variable_identifier_with_digits_and_underscores
# A valid variable identifier may consist of letters, digits, and underscores, provided it does not start with a digit.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
var_123_abc="valid"
[ "$var_123_abc" = "valid" ] || fail "var_123_abc: want 'valid', got [$var_123_abc]"
echo PASS
exit 0
