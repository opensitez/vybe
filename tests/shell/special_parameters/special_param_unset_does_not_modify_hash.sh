#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_unset_does_not_modify_hash
# Executing 'unset #' or 'unset \?' does not remove or modify special parameters.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- "alpha" "beta" "gamma"
unset '#' 2>/dev/null
[ "$#" -eq 3 ] || fail "parameter count altered by unset: want 3, got $#"
unset '?' 2>/dev/null
[ "$?" -eq 0 ] || fail "status altered by unset"
echo PASS
exit 0
