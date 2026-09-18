#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/status_is_determined_after_or_chain
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false || false
s=$?
[ "$s" -eq 1 ] || fail "two failures yields status 1"
true || true
s2=$?
[ "$s2" -eq 0 ] || fail "rhs success should yield status 0"
echo PASS
exit 0
