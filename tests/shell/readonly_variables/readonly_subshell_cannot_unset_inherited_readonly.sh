#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_subshell_cannot_unset_inherited_readonly
# A subshell cannot unset an inherited readonly variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly FIXED_VAL="stay"
(
    unset FIXED_VAL 2>/dev/null
    st=$?
    [ "$st" -ne 0 ] || exit 1
)
st=$?
[ "$st" -eq 0 ] || fail "subshell unexpectedly allowed unsetting inherited readonly variable"
echo PASS
exit 0
