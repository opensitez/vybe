#!/usr/bin/env bash
# vybe-test: bash/arithmetic_base_literals/digit_out_of_range_is_an_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'echo $((08))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "08: want status 1 got $st"
[[ $msg == *"value too great for base"* ]] || fail "08: got [$msg]"
msg=$( eval 'echo $((2#102))' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "2#102: want status 1 got $st"
[[ $msg == *"value too great for base"* ]] || fail "2#102: got [$msg]"
echo PASS
exit 0
