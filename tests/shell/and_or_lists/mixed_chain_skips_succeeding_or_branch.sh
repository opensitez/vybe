#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/mixed_chain_skips_succeeding_or_branch
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
log=""
: || log=or_ran && log=and_after_or
[ "$log" = "and_after_or" ] || fail "or should skip and then and should run"
echo PASS
exit 0
