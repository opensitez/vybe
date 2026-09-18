#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/mixed_chain_skips_failing_and_branch
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
log=""
: && false || log=or_after_fail
[ "$log" = "or_after_fail" ] || fail "failed and branch should trigger or branch"
echo PASS
exit 0
