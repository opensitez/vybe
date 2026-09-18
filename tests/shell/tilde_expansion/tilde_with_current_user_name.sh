#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_with_current_user_name
# The ~username syntax expands to the home directory associated with the specified login user.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
current_user=$(id -un)
[ -n "$current_user" ] || fail "failed to determine current username"
# Evaluate ~<username> dynamically via eval
expanded=$(eval printf '%s' "~$current_user")
# Verify that the expanded directory exists on the system
[ -d "$expanded" ] || fail "~username did not expand to existing directory: got [$expanded]"
echo PASS
exit 0
