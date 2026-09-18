#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/assigning_array_in_arithmetic_expression_is_invalid
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
arr=(1 2)
msg=$( eval 'echo $((arr = 3))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "array assignment should fail"
[[ $msg == *"attempted assignment to non-variable"* ]] || [[ $msg == *"invalid arithmetic"* ]] || fail "msg [$msg]"
echo PASS
exit 0
