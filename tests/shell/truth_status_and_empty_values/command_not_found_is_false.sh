#!/usr/bin/env bash
# vybe-test: bash/truth_status_and_empty_values/command_not_found_is_false
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
if no_such_command_zz 2>/dev/null; then r=t; else r=f; fi
[ "$r" = f ] || fail "127 must be false"
no_such_command_zz 2>/dev/null; st=$?
[ "$st" -eq 127 ] || fail "want 127 got $st"
echo PASS
exit 0
