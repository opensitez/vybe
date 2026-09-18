#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_nameref_aliasing_dynamic_target
# Declaring a nameref where the target name is held in a variable dynamically binds to that target.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
actual_slot="start_value"
chosen_name="actual_slot"
declare -n ref="$chosen_name"
ref="updated_value"
[ "$actual_slot" = "updated_value" ] || fail "dynamic nameref mutation failed: got [$actual_slot]"
echo PASS
exit 0
