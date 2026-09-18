#!/usr/bin/env bash
# vybe-test: bash/readonly_variables/readonly_associative_array_mutation_fails
# An associative array declared with 'readonly -A' rejects key modifications and additions.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
readonly -A LOOKUP=( [user]="admin" [status]="active" )
( LOOKUP[user]="guest" ) 2>/dev/null
st1=$?
[ "$st1" -ne 0 ] || fail "modifying existing key in readonly associative array should fail"

( LOOKUP[new_key]="new_val" ) 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "adding new key to readonly associative array should fail"
echo PASS
exit 0
