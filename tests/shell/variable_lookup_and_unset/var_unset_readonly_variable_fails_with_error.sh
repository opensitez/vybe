#!/usr/bin/env bash
# vybe-test: bash/variable_lookup_and_unset/var_unset_readonly_variable_fails_with_error
# Attempting to unset a variable marked readonly produces a non-zero exit status and error message.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
    readonly frozen="permanent"
    unset frozen 2>/dev/null
    exit $?
)
st=$?
[ "$st" -ne 0 ] || fail "unsetting readonly variable should return non-zero status"
echo PASS
exit 0
