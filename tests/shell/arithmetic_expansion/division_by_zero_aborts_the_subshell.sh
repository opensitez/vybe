#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/division_by_zero_aborts_the_subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'echo $(( 1/0 )); echo unreachable' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "/: want status 1 got $st"
[[ $msg == *"division by 0"* ]] || fail "/: got [$msg]"
[[ $msg != *unreachable* ]] || fail "/: subshell must abort"
msg=$( eval 'echo $(( 5%0 ))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "%: want status 1 got $st"
[[ $msg == *"division by 0"* ]] || fail "%: got [$msg]"
echo PASS
exit 0
