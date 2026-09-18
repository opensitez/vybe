#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_exclamation_initially_unset_or_empty
# Before any background jobs are placed into the background, $! expands to empty.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
fresh_subshell=$( ( printf '%s\n' "$!" ) )
[ -z "$fresh_subshell" ] || fail "\$! was not empty in new subshell: got [$fresh_subshell]"
echo PASS
exit 0
