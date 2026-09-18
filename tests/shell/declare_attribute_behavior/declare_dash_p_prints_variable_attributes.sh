#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_dash_p_prints_variable_attributes
# The 'declare -p var' builtin displays the attributes and assignment string of the specified variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i test_int_prop=42
attr_str=$(declare -p test_int_prop)
case "$attr_str" in
    *"-i"*"test_int_prop="*) : ;;
    *) fail "declare -p output missing -i attribute: got [$attr_str]" ;;
esac
echo PASS
exit 0
