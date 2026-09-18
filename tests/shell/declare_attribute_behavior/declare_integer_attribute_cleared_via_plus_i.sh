#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_integer_attribute_cleared_via_plus_i
# The 'declare +i' flag removes the integer attribute, allowing literal string assignments.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i val=10
declare +i val
val="hello world"
[ "$val" = "hello world" ] || fail "removing integer attribute via +i failed: got [$val]"
echo PASS
exit 0
