#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_empty_variable_remains_readonly
# A variable declared readonly with an empty value remains permanently empty and immutable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly EMPTY_RO=""
[ -z "$EMPTY_RO" ] || fail "empty readonly should be empty"
( EMPTY_RO="filled" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "empty readonly variable should reject assignment"
[ -z "$EMPTY_RO" ] || fail "empty readonly was modified: got [$EMPTY_RO]"
echo PASS
exit 0
