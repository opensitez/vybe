#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_associative_array_key_fallback
# Default value expansions apply cleanly to specific associative array keys.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [existing]="stored_val" )
present_val="${map[existing]:-fallback}"
absent_val="${map[nonexistent]:-default_val}"
[ "$present_val" = "stored_val" ] || fail "present key fallback failed: got [$present_val]"
[ "$absent_val" = "default_val" ] || fail "absent key fallback failed: got [$absent_val]"
echo PASS
exit 0
