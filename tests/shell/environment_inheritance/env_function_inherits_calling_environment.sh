#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_function_inherits_calling_environment
# Shell functions inherit exported environment variables dynamically from the calling environment.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export DYNAMIC_INHERITED="origin"
read_env_fn() {
    [ "$DYNAMIC_INHERITED" = "origin" ] || exit 1
}
read_env_fn
st=$?
[ "$st" -eq 0 ] || fail "function failed to read calling environment variable"
echo PASS
exit 0
