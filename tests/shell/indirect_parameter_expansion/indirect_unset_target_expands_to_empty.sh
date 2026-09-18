#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_unset_target_expands_to_empty
# The ${!ptr} expansion expands to an empty string when the target variable named by ptr is unset.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset nonexistent_target
ptr="nonexistent_target"
val="${!ptr}"
[ -z "$val" ] || fail "unset target should expand to empty: got [$val]"
echo PASS
exit 0
