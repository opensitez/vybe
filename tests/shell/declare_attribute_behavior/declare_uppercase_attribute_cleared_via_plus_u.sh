#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_uppercase_attribute_cleared_via_plus_u
# The 'declare +u' flag clears the uppercase attribute, preserving subsequent lowercase assignments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -u tag="hello"
declare +u tag
tag="lowercase_now"
[ "$tag" = "lowercase_now" ] || fail "+u failed to clear uppercase attribute: got [$tag]"
echo PASS
exit 0
