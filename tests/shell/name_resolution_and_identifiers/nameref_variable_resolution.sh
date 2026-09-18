#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/nameref_variable_resolution
# A nameref variable declared with 'declare -n' resolves assignments and reads to the referenced identifier.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
target_var="initial"
declare -n ref=target_var
[ "$ref" = "initial" ] || fail "nameref read: want 'initial', got [$ref]"
ref="modified_via_ref"
[ "$target_var" = "modified_via_ref" ] || fail "nameref write target: want 'modified_via_ref', got [$target_var]"
echo PASS
exit 0
