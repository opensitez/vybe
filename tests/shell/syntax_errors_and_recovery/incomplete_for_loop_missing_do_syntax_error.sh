#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/incomplete_for_loop_missing_do_syntax_error
# A for loop missing the required 'do' keyword produces a syntax error (status 2).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'for x in 1 2; done' 2>/dev/null
st=$?
[ "$st" -eq 2 ] || fail "for loop missing do: want status 2, got $st"
echo PASS
exit 0
