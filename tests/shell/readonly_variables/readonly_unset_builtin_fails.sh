#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_unset_builtin_fails
# Attempting to unset a readonly variable fails with an error and non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly INDELIBLE="permanent"
( unset INDELIBLE ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unsetting readonly variable should return non-zero exit code"
[ "$INDELIBLE" = "permanent" ] || fail "readonly variable was removed: got [$INDELIBLE]"
echo PASS
exit 0
