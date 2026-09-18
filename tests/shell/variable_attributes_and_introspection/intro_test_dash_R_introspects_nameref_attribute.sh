#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_test_dash_R_introspects_nameref_attribute
# The [[ -R var ]] conditional operator explicitly introspects whether var is a nameref.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
core="val"
declare -n ref=core
[[ -R ref ]] || fail "[[ -R ref ]] should evaluate to true for nameref"
[[ ! -R core ]] || fail "[[ -R core ]] should evaluate to false for standard variable"
echo PASS
exit 0
