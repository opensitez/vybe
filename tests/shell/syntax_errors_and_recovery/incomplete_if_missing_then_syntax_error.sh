#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/incomplete_if_missing_then_syntax_error
# An if compound command missing the required 'then' keyword produces an unexpected token syntax error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'if true; fi' 2>/dev/null
st=$?
[ "$st" -eq 2 ] || fail "if true; fi missing then: want status 2, got $st"
echo PASS
exit 0
