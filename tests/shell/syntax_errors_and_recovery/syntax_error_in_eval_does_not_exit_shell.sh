#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/syntax_error_in_eval_does_not_exit_shell
# A non-interactive, non-posix shell survives a syntax error inside eval.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'if true; then' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "eval must report failure"
survived=yes
[ "$survived" = yes ] || fail "unreachable"
echo PASS
exit 0
