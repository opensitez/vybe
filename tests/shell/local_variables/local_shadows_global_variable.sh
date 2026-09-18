#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_shadows_global_variable
# Declaring a local variable inside a function shadows an outer variable of the same name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="global_value"
test_fn() {
    local x="local_value"
    [ "$x" = "local_value" ] || exit 1
}
test_fn
st=$?
[ "$st" -eq 0 ] || fail "local variable failed to shadow global"
echo PASS
exit 0
