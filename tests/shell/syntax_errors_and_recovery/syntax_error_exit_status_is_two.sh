#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/syntax_error_exit_status_is_two
# A shell that hits a syntax error in its script exits with status 2.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
"$BASH" -c 'echo a ) b' 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "want 2 got $st"
echo PASS
exit 0
