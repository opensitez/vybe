#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_child_modifications_do_not_leak_to_parent
# Changes made to inherited environment variables by a child process do not reflect back into the parent.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export SAFE_PARENT_VAR="stable_parent_state"
"$BASH" -c 'SAFE_PARENT_VAR="child_overwrite"; export SAFE_PARENT_VAR'
[ "$SAFE_PARENT_VAR" = "stable_parent_state" ] || fail "parent environment tainted by child execution: got [$SAFE_PARENT_VAR]"
echo PASS
exit 0
