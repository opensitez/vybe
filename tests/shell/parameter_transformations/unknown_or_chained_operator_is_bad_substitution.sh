#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/unknown_or_chained_operator_is_bad_substitution
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=abc
msg=$( eval 'echo "${x@Z}"' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "unknown operator: want status 1 got $st"
[[ $msg == *"bad substitution"* ]] || fail "unknown operator: got [$msg]"
msg=$( eval 'echo "${x@Q@Q}"' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "chained: want status 1 got $st"
[[ $msg == *"bad substitution"* ]] || fail "chained: got [$msg]"
echo PASS
exit 0
