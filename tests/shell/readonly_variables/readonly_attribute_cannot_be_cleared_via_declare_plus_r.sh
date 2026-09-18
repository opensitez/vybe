#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_attribute_cannot_be_cleared_via_declare_plus_r
# The readonly attribute cannot be unset using 'declare +r'; the shell rejects the operation.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly LOCKED="secure"
( declare +r LOCKED ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "declare +r on readonly variable should return non-zero exit status"
echo PASS
exit 0
