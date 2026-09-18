#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_array_keys_iteration_in_for_loop
# Iterating over "${!map[@]}" traverses the keys of an associative array cleanly in a for loop.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A metrics=( [cpu]="80" [mem]="60" [disk]="45" )
audit=""
for key in "${!metrics[@]}"; do
    audit+="$key:${metrics[$key]};"
done
case "$audit" in
    *"cpu:80"*|*"mem:60"*|*"disk:45"*) : ;;
    *) fail "associative array key iteration corrupted: got [$audit]" ;;
esac
echo PASS
exit 0
