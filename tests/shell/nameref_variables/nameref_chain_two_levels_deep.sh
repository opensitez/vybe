#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_chain_two_levels_deep
# Chaining namerefs (ref1 -> ref2 -> actual) resolves reads and writes to the final target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
core_target="deep_payload"
declare -n middle_ref=core_target
declare -n top_ref=middle_ref
[ "$top_ref" = "deep_payload" ] || fail "chained nameref read failed: got [$top_ref]"
top_ref="overwritten_from_top"
[ "$core_target" = "overwritten_from_top" ] || fail "chained nameref mutation failed: got [$core_target]"
echo PASS
exit 0
