#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/self_reference_after_update_causes_recursion_guard
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=1
x='x + 1'
msg=$( eval 'echo $((x))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "self-referential re-assignment must fail"
[[ $msg == *"recursion level exceeded"* ]] || fail "unexpected: [$msg]"
echo PASS
exit 0
