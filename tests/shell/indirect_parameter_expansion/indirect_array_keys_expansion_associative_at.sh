#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_array_keys_expansion_associative_at
# The ${!map[@]} expansion expands to the keys of an associative array.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [alpha]="1" [beta]="2" )
keys="${!map[@]}"
case "$keys" in
    *"alpha"*"beta"*|*"beta"*"alpha"*) : ;;
    *) fail "associative array keys missing: got [$keys]" ;;
esac
echo PASS
exit 0
