#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_defaults_to_local_inside_function
# Without the -g flag, 'declare' inside a function behaves like 'local', scoping the variable to the function.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset LOCAL_BY_DEFAULT
local_creator_fn() {
    declare LOCAL_BY_DEFAULT="scoped_val"
    [ "$LOCAL_BY_DEFAULT" = "scoped_val" ] || exit 1
}
local_creator_fn
st=$?
[ "$st" -eq 0 ] || fail "local execution failed"
[[ ! -v LOCAL_BY_DEFAULT ]] || fail "declare inside function leaked to global scope without -g"
echo PASS
exit 0
