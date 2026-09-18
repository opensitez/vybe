#!/usr/bin/env bash
# vybe-test: bash/background_list_status/wait_and_report_stderr_for_bad_pid
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(wait 999999 2>&1); st=$?
[ "$st" -eq 1 ] || fail "bad pid should fail"
[ -n "$msg" ] || fail "should emit error"
echo PASS
exit 0
