#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_nameref_retargeting_to_another_variable
# Re-declaring 'declare -n ref=new_target' points the nameref to a different target variable.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
varA="initial_A"
varB="initial_B"
declare -n pointer=varA
pointer="mutated_A"
[ "$varA" = "mutated_A" ] || fail "first target mutation failed"

declare -n pointer=varB
pointer="mutated_B"
[ "$varB" = "mutated_B" ] || fail "retargeted mutation failed"
[ "$varA" = "mutated_A" ] || fail "varA was modified after retargeting"
echo PASS
exit 0
