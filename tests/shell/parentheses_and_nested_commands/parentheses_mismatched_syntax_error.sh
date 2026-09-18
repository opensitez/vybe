#!/usr/bin/env bash
# vybe-test: bash/parentheses_and_nested_commands/parentheses_mismatched_syntax_error
# Mismatched or unclosed parentheses trigger a syntax error (status 2).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'echo $(( (1 + 2) * 3' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unclosed arithmetic parenthesis should fail syntax parsing"

eval '( echo unmatched' 2>/dev/null
st2=$?
[ "$st2" -ne 0 ] || fail "unclosed subshell parenthesis should fail syntax parsing"
echo PASS
exit 0
