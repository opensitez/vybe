#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_name_constructed_from_command_substitution
# Dynamic variable identifiers can be constructed using command substitution output $( ... ).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
suffix=$(printf 'runtime_token\n')
printf -v "prefix_${suffix}" "%s" "payload"
ptr="prefix_runtime_token"
[ "${!ptr}" = "payload" ] || fail "cmdsub-constructed variable name failed: got [${!ptr}]"
echo PASS
exit 0
