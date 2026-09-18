#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unmatched_brace_parameter_expansion_bad_substitution
# An unclosed parameter expansion brace ${var triggers a bad substitution syntax error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'echo ${unclosed_var' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unclosed parameter brace should return non-zero syntax error"
echo PASS
exit 0
