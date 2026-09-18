#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unbalanced_double_bracket_is_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval '[[ a == a' 2>&1); st=$?
[ "$st" -ne 0 ] || fail "missing ]] must fail"
[[ $msg == *"syntax error"* ]] || fail "got [$msg]"
echo PASS
exit 0
