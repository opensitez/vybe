#!/usr/bin/env bash
# vybe-test: bash/arithmetic_increment_and_assignment/assignment_to_non_variable_is_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'echo $(( 5 = 3 ))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"attempted assignment to non-variable"* ]] || fail "got [$msg]"
msg=$( eval 'echo $(( 5++ ))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "5++: want status 1 got $st"
echo PASS
exit 0
