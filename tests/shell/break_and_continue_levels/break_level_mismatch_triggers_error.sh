#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/break_level_mismatch_triggers_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(break 2 2>/dev/null)
[ "$?" -ne 0 ] || fail "break 2 outside loops should error"
[ -n "$msg" ] || fail "error message expected"
echo PASS
exit 0
