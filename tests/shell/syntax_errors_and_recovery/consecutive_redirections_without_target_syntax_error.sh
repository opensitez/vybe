#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/consecutive_redirections_without_target_syntax_error
# Providing consecutive redirection operators without a filename target for the first is a syntax error.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval 'cat > < /dev/null' 2>/dev/null
st=$?
[ "$st" -eq 2 ] || fail "consecutive redirections: want status 2, got $st"
echo PASS
exit 0
