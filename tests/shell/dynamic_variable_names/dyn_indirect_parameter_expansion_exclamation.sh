#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_indirect_parameter_expansion_exclamation
# The syntax ${!pointer} expands to the value of the variable named by pointer.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target_var="target_payload"
pointer="target_var"
[ "${!pointer}" = "target_payload" ] || fail "indirect expansion failed: got [${!pointer}]"
echo PASS
exit 0
