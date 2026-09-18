#!/usr/bin/env bash
# vybe-test: bash/nameref_variables/nameref_assignment_mutates_underlying_target
# Assigning a value to a nameref mutates the underlying referenced target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source_target="initial_val"
declare -n ref=source_target
ref="mutated_via_ref"
[ "$source_target" = "mutated_via_ref" ] || fail "target variable not mutated: got [$source_target]"
[ "$ref" = "mutated_via_ref" ] || fail "nameref value mismatch"
echo PASS
exit 0
