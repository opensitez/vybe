#!/usr/bin/env bash
# vybe-test: bash/syntax_errors_and_recovery/unterminated_quote_reports_unexpected_eof
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(eval 'echo "abc' 2>&1); st=$?
[ "$st" -ne 0 ] || fail "unterminated quote must fail"
[[ $msg == *"unexpected EOF"* ]] || fail "want unexpected EOF diagnostic, got [$msg]"
echo PASS
exit 0
