#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/syntax_error_reports_line_number
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$("$BASH" -c $'echo ok\necho ok\necho a ) b' 2>&1)
[[ $msg == *"line 3:"* ]] || fail "want line 3 in diagnostic got [$msg]"
echo PASS
exit 0
