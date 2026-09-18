#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/grouping_prevents_chain_spillover
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
log=""
(false && : ) || log=outer
[ "$log" = "outer" ] || fail "outer or must run"
log=""
(false || : ) && log=outer
[ "$log" = "" ] || fail "outer and must be skipped"
echo PASS
exit 0
