#!/usr/bin/env bash
# vybe-test: bash/break_and_continue_levels/continue_level_mismatch_triggers_error
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$(continue 3 2>/dev/null)
[ "$?" -ne 0 ] || fail "continue 3 outside loops should fail"
[ -n "$msg" ] || fail "expected error"
echo PASS
exit 0
