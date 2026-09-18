#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/assigning_to_array_index_from_lookup
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=(9 8)
idx=1
(( a[idx] = 15 ))
[ "${a[1]}" -eq 15 ] || fail "a[idx] assigned"
[ $((a[idx])) -eq 15 ] || fail "a[idx] read as number"
echo PASS
exit 0
