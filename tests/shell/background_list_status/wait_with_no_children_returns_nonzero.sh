#!/usr/bin/env bash
# vybe-test: bash/background_list_status/wait_with_no_children_returns_nonzero
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
wait 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "wait with no children should be non-zero"
echo PASS
exit 0
