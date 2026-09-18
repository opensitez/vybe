#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_declare_dash_p_introspects_nameref
# Running 'declare -p' on a nameref variable displays the -n flag and points to target variable name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source_target="data"
declare -n ref=source_target
output=$(declare -p ref)
case "$output" in
    *'declare -n ref="source_target"'*) : ;;
    *) fail "declare -p output missing nameref info: got [$output]" ;;
esac
echo PASS
exit 0
