#!/usr/bin/env bash
# vybe-test: bash/parameter_defaults/param_default_indexed_array_element_fallback
# Default value expansions apply cleanly to specific indexed array elements.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=( "present_elem" )
elem0="${arr[0]:-fallback0}"
elem5="${arr[5]:-fallback5}"
[ "$elem0" = "present_elem" ] || fail "elem0 fallback failed: got [$elem0]"
[ "$elem5" = "fallback5" ] || fail "elem5 fallback failed: got [$elem5]"
echo PASS
exit 0
