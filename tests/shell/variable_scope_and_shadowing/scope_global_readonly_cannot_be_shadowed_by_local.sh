#!/usr/bin/env bash
# vybe-test: bash/variable_scope_and_shadowing/scope_global_readonly_cannot_be_shadowed_by_local
# In Bash, attempting to shadow a global readonly variable with 'local var' fails with an error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly IMMUTABLE_GLOBAL="frozen"
shadow_readonly_fn() {
    local IMMUTABLE_GLOBAL="shadow_attempt" 2>/dev/null
    st=$?
    [ "$st" -ne 0 ] || fail "local shadowing of global readonly should return non-zero exit status"
}
shadow_readonly_fn
[ "$IMMUTABLE_GLOBAL" = "frozen" ] || fail "global readonly variable altered: got [$IMMUTABLE_GLOBAL]"
echo PASS
exit 0
