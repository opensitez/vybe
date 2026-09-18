#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_lowercase_attribute_converts_on_assignment
# The 'declare -l' flag converts all assigned text automatically to lowercase characters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -l lower_var
lower_var="MixedCase_STRING_123"
[ "$lower_var" = "mixedcase_string_123" ] || fail "lowercase conversion failed: got [$lower_var]"
echo PASS
exit 0
