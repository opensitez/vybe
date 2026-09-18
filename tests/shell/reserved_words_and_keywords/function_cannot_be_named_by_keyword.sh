#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/function_cannot_be_named_by_keyword
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval 'if() { :; }' 2>&1 ); st=$?
[ "$st" -ne 0 ] || fail "defining a function named if must fail"
[[ $msg == *"syntax error"* ]] || fail "want a syntax error diagnostic, got [$msg]"
echo PASS
exit 0
