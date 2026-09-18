#!/usr/bin/env bash
# vybe-test: bash/indirect_parameter_expansion/indirect_associative_array_key_dereference
# The ${!ptr} expansion where ptr="map[key]" dereferences the associative array element at key.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
declare -A map=( [username]="admin_user" [protocol]="https" )
ptr="map[username]"
[ "${!ptr}" = "admin_user" ] || fail "indirect associative array access failed: got [${!ptr}]"
ptr="map[protocol]"
[ "${!ptr}" = "https" ] || fail "indirect associative array protocol failed: got [${!ptr}]"
echo PASS
exit 0
