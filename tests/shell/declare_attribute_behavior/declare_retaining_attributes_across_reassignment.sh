#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_retaining_attributes_across_reassignment
# Attributes set via 'declare' persist across multiple standard variable reassignments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -l sticky_lower="FIRST"
[ "$sticky_lower" = "first" ] || fail "initial lowercase failed"
sticky_lower="SECOND"
[ "$sticky_lower" = "second" ] || fail "subsequent assignment failed to enforce lowercase"
sticky_lower="THIRD"
[ "$sticky_lower" = "third" ] || fail "third assignment failed to enforce lowercase"
echo PASS
exit 0
