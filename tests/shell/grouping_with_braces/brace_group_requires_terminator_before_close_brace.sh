#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_requires_terminator_before_close_brace
# The closing brace '}' requires a preceding command terminator (semicolon or newline).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '{ echo 1 }' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "brace group without terminator before '}' should produce syntax error"
echo PASS
exit 0
