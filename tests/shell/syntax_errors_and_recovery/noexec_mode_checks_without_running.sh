#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/noexec_mode_checks_without_running
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$("$BASH" -n -c 'echo ran; exit 3' 2>&1); st=$?
[ "$st" -eq 0 ] || fail "valid script under -n must exit 0, got $st"
[ -z "$out" ] || fail "-n must not execute, got [$out]"
"$BASH" -n -c 'if true; then' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "-n must still report a syntax error"
echo PASS
exit 0
