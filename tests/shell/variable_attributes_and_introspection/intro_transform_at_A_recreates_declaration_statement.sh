#!/usr/bin/env bash
# vybe-test: bash/variable_attributes_and_introspection/intro_transform_at_A_recreates_declaration_statement
# The ${var@A} parameter transformation expands to an assignment statement recreating the variable with attributes.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -i base_number=77
stmt="${base_number@A}"
case "$stmt" in
    *"declare -i base_number="*) : ;;
    *) fail "\${base_number@A} did not recreate declare statement: got [$stmt]" ;;
esac
echo PASS
exit 0
