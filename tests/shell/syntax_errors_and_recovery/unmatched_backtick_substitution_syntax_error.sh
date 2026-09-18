#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unmatched_backtick_substitution_syntax_error
# An unmatched backtick in command substitution triggers an unexpected EOF syntax error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'echo `echo unclosed' 2>/dev/null
st=$?
[ "$st" -eq 2 ] || fail "unclosed backtick: want status 2, got $st"
echo PASS
exit 0
