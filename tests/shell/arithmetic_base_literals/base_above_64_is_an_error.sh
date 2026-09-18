#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/base_above_64_is_an_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'echo $((65#1))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"invalid arithmetic base"* ]] || fail "got [$msg]"
[ $((64#10)) -eq 64 ] || fail "64 is the largest valid base"
echo PASS
exit 0
