#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_declare_dash_p_introspects_associative_array
# Running 'declare -p' on an associative array displays key-value pairs and the -A flag.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A user_map=( [admin]="superuser" )
output=$(declare -p user_map)
case "$output" in
    *"declare -A user_map="*'[admin]="superuser"'*) : ;;
    *) fail "declare -p output missing associative array details: got [$output]" ;;
esac
echo PASS
exit 0
