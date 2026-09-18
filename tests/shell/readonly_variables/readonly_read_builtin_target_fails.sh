#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_read_builtin_target_fails
# Attempting to assign into a readonly variable via the 'read' builtin fails with non-zero exit status.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly TARGET="protected"
( read -r TARGET <<< "incoming" ) 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "read into readonly variable should fail with non-zero exit code"
[ "$TARGET" = "protected" ] || fail "readonly variable was altered: got [$TARGET]"
echo PASS
exit 0
