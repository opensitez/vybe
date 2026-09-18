#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_hyphen_reflects_active_shell_flags
# The $- parameter expands to the current option flags specified upon invocation or via set.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
flags="$-"
[ -n "$flags" ] || fail "\$- option string should not be empty"
echo PASS
exit 0
