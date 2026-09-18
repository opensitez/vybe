#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_child_inherits_exported_variable
# Child shell processes inherit exported environment variables from the parent shell.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export ACTIVE_ENV_KEY="inherited_secret"
child_val=$( "$BASH" -c 'printf "%s\n" "$ACTIVE_ENV_KEY"' )
[ "$child_val" = "inherited_secret" ] || fail "child process failed to inherit exported variable: got [$child_val]"
echo PASS
exit 0
