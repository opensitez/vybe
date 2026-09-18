#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_unassigned_declaration_is_unset_initially
# Declaring 'local var' without assignment creates a local slot that shadows outer but is initially unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
v="outer_value"
inspect_local() {
    local v
    [[ ! -v v ]] || fail "unassigned local v should be unset (-v)"
    [ -z "$v" ] || fail "unassigned local v should expand to empty"
    v="now_assigned"
    [ "$v" = "now_assigned" ] || fail "local assignment failed"
}
inspect_local
[ "$v" = "outer_value" ] || fail "outer variable altered: got [$v]"
echo PASS
exit 0
