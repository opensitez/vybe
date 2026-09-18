#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_nameref_declaration_dash_n
# Declaring 'local -n ref=target' creates a local nameref referencing an outer variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
external_data="original_payload"
mutate_ref() {
    local -n ref_var=$1
    ref_var="modified_payload"
}
mutate_ref external_data
[ "$external_data" = "modified_payload" ] || fail "local nameref failed to mutate external target"
[[ ! -v ref_var ]] || fail "local nameref ref_var should not exist in outer scope"
echo PASS
exit 0
