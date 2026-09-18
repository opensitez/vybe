#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_a_returns_lowercase_and_uppercase_attribute
# The ${var@a} transformation contains 'l' for lowercase and 'u' for uppercase variables.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -l low="TEXT"
declare -u upp="text"
attrs_low="${low@a}"
attrs_upp="${upp@a}"
case "$attrs_low" in
    *l*) : ;;
    *) fail "lowercase attribute missing: got [$attrs_low]" ;;
esac
case "$attrs_upp" in
    *u*) : ;;
    *) fail "uppercase attribute missing: got [$attrs_upp]" ;;
esac
echo PASS
exit 0
