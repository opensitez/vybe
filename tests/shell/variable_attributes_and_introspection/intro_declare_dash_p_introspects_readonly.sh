#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_declare_dash_p_introspects_readonly
# Running 'declare -p' on a readonly variable outputs declaration syntax including the -r flag.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly SYSTEM_UUID="uuid_998877"
output=$(declare -p SYSTEM_UUID)
case "$output" in
    *"declare -r SYSTEM_UUID="*) : ;;
    *) fail "declare -p output missing -r attribute: got [$output]" ;;
esac
echo PASS
exit 0
