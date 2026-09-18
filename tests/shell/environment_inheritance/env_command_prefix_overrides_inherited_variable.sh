#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_command_prefix_overrides_inherited_variable
# Prepending an assignment to a command invocation overrides any inherited value for that command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export TARGET_CONFIG="default_config"
override_val=$( TARGET_CONFIG="overridden_config" "$BASH" -c 'printf "%s\n" "$TARGET_CONFIG"' )
[ "$override_val" = "overridden_config" ] || fail "prefix override failed: got [$override_val]"
[ "$TARGET_CONFIG" = "default_config" ] || fail "parent environment variable was modified: got [$TARGET_CONFIG]"
echo PASS
exit 0
