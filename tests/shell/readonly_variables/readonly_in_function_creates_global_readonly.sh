#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_in_function_creates_global_readonly
# Using 'readonly' (without 'local') inside a function creates a global readonly variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare_ro_fn() {
    readonly GLOBAL_RO="fn_created"
}
declare_ro_fn
[ "$GLOBAL_RO" = "fn_created" ] || fail "global readonly not visible outside function"
( GLOBAL_RO="altered" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "function-created readonly should remain immutable outside function"
echo PASS
exit 0
