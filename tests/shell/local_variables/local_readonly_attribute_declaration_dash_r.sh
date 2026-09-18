#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_readonly_attribute_declaration_dash_r
# Declaring 'local -r ro_var=val' prevents mutation or unsetting of the variable within the local scope.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly_fn() {
    local -r frozen="read_only_val"
    ( frozen="attempted_write" ) 2>/dev/null
    st=$?
    [ "$st" -ne 0 ] || fail "mutation of local -r variable should fail"
}
readonly_fn
echo PASS
exit 0
