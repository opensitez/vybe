#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/self_reference_exceeds_recursion_level
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=x
msg=$( eval 'echo $((x))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"recursion level exceeded"* ]] || fail "got [$msg]"
echo PASS
exit 0
