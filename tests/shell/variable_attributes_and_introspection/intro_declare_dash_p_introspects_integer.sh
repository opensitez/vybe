#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_declare_dash_p_introspects_integer
# Running 'declare -p' on an integer variable outputs declaration syntax including the -i flag.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i target_metric=123
output=$(declare -p target_metric)
case "$output" in
    *"declare -i target_metric="*) : ;;
    *) fail "declare -p output missing -i attribute: got [$output]" ;;
esac
echo PASS
exit 0
