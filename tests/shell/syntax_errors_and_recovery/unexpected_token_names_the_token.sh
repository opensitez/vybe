#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unexpected_token_names_the_token
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'echo a ) b' 2>&1); st=$?
[ "$st" -ne 0 ] || fail "stray ) must fail"
[[ $msg == *"syntax error near unexpected token"*")"* ]] || fail "got [$msg]"
echo PASS
exit 0
