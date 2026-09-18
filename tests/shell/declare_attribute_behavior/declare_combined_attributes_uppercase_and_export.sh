#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_combined_attributes_uppercase_and_export
# Attributes can be combined in a single flag (e.g. 'declare -ux' for uppercase and export).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -ux DUAL_FLAG="lowercase_value"
[ "$DUAL_FLAG" = "LOWERCASE_VALUE" ] || fail "uppercase attribute failed in combined flag: got [$DUAL_FLAG]"
res=$(
    "$BASH" -c 'printf "%s\n" "$DUAL_FLAG"'
)
[ "$res" = "LOWERCASE_VALUE" ] || fail "export attribute failed in combined flag: got [$res]"
echo PASS
exit 0
