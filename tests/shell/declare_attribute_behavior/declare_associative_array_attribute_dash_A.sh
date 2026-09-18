#!/usr/bin/env bash
# vybe-test: bash/declare_attribute_behavior/declare_associative_array_attribute_dash_A
# The 'declare -A' flag declares an associative array mapping string keys to values.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map
map["first_name"]="John"
map["last_name"]="Doe"
[ "${map["first_name"]}" = "John" ] || fail "map first_name failed"
[ "${map["last_name"]}" = "Doe" ] || fail "map last_name failed"
[ "${#map[@]}" -eq 2 ] || fail "map length: want 2, got ${#map[@]}"
echo PASS
exit 0
