#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/source_syntax_error_returns_two_and_recovers
# A syntax error in a sourced file causes the source command to return status 2, and the caller recovers.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
source <(printf 'if true; then\n') 2>/dev/null
st=$?
[ "$st" -eq 2 ] || fail "source syntax error status: want 2, got $st"
echo PASS
exit 0
