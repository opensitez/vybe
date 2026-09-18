#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_clean_invocation_with_env_dash_i
# Invoking a command with 'env -i' clears the inherited environment completely.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export SENTINEL_DATA="should_be_cleared"
saw_sentinel=$( env -i "$BASH" -c 'printf "%s\n" "$SENTINEL_DATA"' )
[ -z "$saw_sentinel" ] || fail "env -i failed to clear inherited environment: saw [$saw_sentinel]"
echo PASS
exit 0
