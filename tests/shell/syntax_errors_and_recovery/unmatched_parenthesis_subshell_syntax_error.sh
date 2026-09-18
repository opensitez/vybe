#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unmatched_parenthesis_subshell_syntax_error
# An open parenthesis without a matching close parenthesis causes a parse error (status 2).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '( echo unclosed' 2>/dev/null
st=$?
[ "$st" -eq 2 ] || fail "unclosed subshell: want status 2, got $st"
echo PASS
exit 0
