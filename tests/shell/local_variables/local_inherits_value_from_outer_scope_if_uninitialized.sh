#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_inherits_value_from_outer_scope_if_uninitialized
# In Bash, an uninitialized 'local var' shadows the outer variable by creating an unset local variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="outer_val"
test_shadow_unset() {
    local x
    # In Bash, local x without assignment shadows outer and is initially unset
    [[ ! -v x ]] || fail "unassigned local x should be unset"
    [ -z "$x" ] || fail "unassigned local x should be empty"
}
test_shadow_unset
[ "$x" = "outer_val" ] || fail "outer variable altered: got [$x]"
echo PASS
exit 0
