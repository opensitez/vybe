#!/usr/bin/env bash
# vybe-test: bash/dynamic_variable_names/dyn_indirect_associative_array_element_access
# Indirect expansion ${!pointer} where pointer="map[key]" retrieves the element with that string key.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A users=( [admin]="superuser" [guest]="anonymous" )
key="admin"
ptr="users[$key]"
[ "${!ptr}" = "superuser" ] || fail "indirect associative array access failed: got [${!ptr}]"
echo PASS
exit 0
