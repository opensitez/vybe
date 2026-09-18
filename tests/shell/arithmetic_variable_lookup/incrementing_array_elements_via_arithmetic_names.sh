#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/incrementing_array_elements_via_arithmetic_names
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=(5 10 15)
i=0
(( a[i] ))
[ "${a[0]}" -eq 5 ] || fail "a[0] should be 5"
(( a[i]++ ))
[ "${a[0]}" -eq 6 ] || fail "a[0] incremented to 6"
i=1
(( a[i]-- ))
[ "${a[1]}" -eq 9 ] || fail "a[1] decremented to 9"
(( a[i+1]++ ))
[ "${a[2]}" -eq 16 ] || fail "a[2] incremented to 16"
echo PASS
exit 0
