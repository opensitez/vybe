#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_exclamation_expands_to_target_name
# When applied to a nameref variable, ${!ref} expands to the name of the target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
underlying_storage="payload"
declare -n ref=underlying_storage
target_name="${!ref}"
[ "$target_name" = "underlying_storage" ] || fail "\${!ref} failed: want 'underlying_storage', got [$target_name]"
echo PASS
exit 0
