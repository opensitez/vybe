#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/function_shadows_builtin_by_default
# A defined shell function takes precedence over a shell builtin with the same name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
pwd() { printf 'shadowed_pwd\n'; }
res=$(pwd)
unset -f pwd
[ "$res" = "shadowed_pwd" ] || fail "function did not shadow builtin: got [$res]"
echo PASS
exit 0
