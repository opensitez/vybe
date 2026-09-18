#!/usr/bin/env bash
# vybe-test: bash/arithmetic_expansion/negative_exponent_is_an_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'echo $(( 2 ** -1 ))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "want status 1 got $st"
[[ $msg == *"exponent less than 0"* ]] || fail "got [$msg]"
echo PASS
exit 0
