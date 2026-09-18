#!/usr/bin/env bash
# vybe-test: bash/parameter_expansion_basics/malformed_braces_are_bad_substitution
# Nested ${${x}} and empty ${} are "bad substitution": an expansion error
# that aborts the (sub)shell with status 1. (A space after ${ is not covered
# here: since bash 5.3 "${ cmd; }" is the no-fork command substitution.)
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=y
msg=$( eval 'echo ${${x}}' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "nested: want status 1 got $st"
[[ $msg == *"bad substitution"* ]] || fail "nested: got [$msg]"
msg=$( eval 'echo ${}' 2>&1 ); st=$?
[ "$st" -eq 1 ] || fail "empty: want status 1 got $st"
[[ $msg == *"bad substitution"* ]] || fail "empty: got [$msg]"
msg=$( eval 'echo ${x-' 2>&1 ); st=$?
[ "$st" -ne 0 ] || fail "unterminated: must fail"
echo PASS
exit 0
