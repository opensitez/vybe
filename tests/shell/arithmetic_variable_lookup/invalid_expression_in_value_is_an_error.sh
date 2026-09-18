#!/usr/bin/env bash
# vybe-test: bash/arithmetic_variable_lookup/invalid_expression_in_value_is_an_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='1+'
msg=$( eval 'echo $((x))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"operand expected"* ]] || fail "got [$msg]"
echo PASS
exit 0
