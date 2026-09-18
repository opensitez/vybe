#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/redirection_without_target_is_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'echo hi >' 2>&1); st=$?
[ "$st" -ne 0 ] || fail "must fail"
[[ $msg == *"unexpected token"*newline* ]] || fail "want unexpected token newline got [$msg]"
echo PASS
exit 0
