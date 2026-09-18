#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_inherited_by_subshell_remains_immutable
# Readonly variables retain their readonly attribute when inherited into a subshell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly IMMUTABLE="subshell_test"
(
    ( IMMUTABLE="modified_in_sub" ) 2>/dev/null
    st=$?
    [ "$st" -ne 0 ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "readonly variable was modifiable inside subshell"
echo PASS
exit 0
