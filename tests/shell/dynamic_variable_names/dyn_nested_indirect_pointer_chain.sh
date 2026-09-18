#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_nested_indirect_pointer_chain
# Multi-level indirect variable resolution dereferences sequential pointers step by step.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
final_destination="treasure"
ptr_b="final_destination"
ptr_a="ptr_b"
# Step 1: dereference ptr_a to get name held in ptr_b
step1="${!ptr_a}"
[ "$step1" = "final_destination" ] || fail "step 1 indirect resolution failed: got [$step1]"
# Step 2: dereference step1 to get value in final_destination
step2="${!step1}"
[ "$step2" = "treasure" ] || fail "step 2 indirect resolution failed: got [$step2]"
echo PASS
exit 0
