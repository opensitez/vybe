#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_lowercase_attribute_cleared_via_plus_l
# The 'declare +l' flag clears the lowercase attribute, preserving subsequent mixed-case assignments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -l text="INITIAL"
declare +l text
text="MixedCase"
[ "$text" = "MixedCase" ] || fail "+l failed to clear lowercase attribute: got [$text]"
echo PASS
exit 0
