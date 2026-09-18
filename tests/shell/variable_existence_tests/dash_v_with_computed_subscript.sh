#!/usr/bin/env bash
# vybe-test: bash/variable_existence_tests/dash_v_with_computed_subscript
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
a=(1 2 3); i=2
[[ -v a[i] ]] || fail "arithmetic subscript"
[[ -v a[i+1] ]] && fail "a[3] does not exist"
[[ -v "a[$i]" ]] || fail "quoted expanded subscript"
declare -A m=([k]=v); key=k
[[ -v m[$key] ]] || fail "associative key from variable"
[[ -v m[other] ]] && fail "missing associative key"
echo PASS
exit 0
