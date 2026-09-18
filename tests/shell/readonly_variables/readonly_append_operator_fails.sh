#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_append_operator_fails
# Attempting to use '+=' append operator on a readonly variable fails with an error and non-zero exit code.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly STEM="base"
( STEM+="_suffix" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "+= on readonly variable should fail"
[ "$STEM" = "base" ] || fail "readonly variable value modified: got [$STEM]"
echo PASS
exit 0
