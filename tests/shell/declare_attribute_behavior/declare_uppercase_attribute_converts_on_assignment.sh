#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_uppercase_attribute_converts_on_assignment
# The 'declare -u' flag converts all assigned text automatically to uppercase characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -u upper_var
upper_var="mixed_lower_text"
[ "$upper_var" = "MIXED_LOWER_TEXT" ] || fail "uppercase conversion failed: got [$upper_var]"
echo PASS
exit 0
