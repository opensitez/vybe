#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/empty_array_name_as_index_not_numeric
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=([ ]=5)
name=''
msg=$( eval 'echo $((a[name]))' 2>&1 ); st=$?
[ "$st" -eq 0 ] || fail "empty name index should parse as 0 index"
[ "$((a[name]))" -eq 5 ] || fail "a[''] should map to a[0]"
echo PASS
