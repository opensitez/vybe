#!/usr/bin/env bash
# vybe-test: bash/name_resolution_and_identifiers/alias_shadows_function_and_builtin_when_enabled
# When alias expansion is enabled, an alias takes precedence over functions and builtins.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
shopt -s expand_aliases
test_target() { printf 'function_val\n'; }
alias test_target="printf 'alias_val\n'"
out=$(test_target)
unalias test_target
[ "$out" = "alias_val" ] || fail "alias did not shadow function: got [$out]"
echo PASS
exit 0
