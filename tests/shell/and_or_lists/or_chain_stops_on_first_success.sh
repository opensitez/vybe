#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/or_chain_stops_on_first_success
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
seen=none
: || seen=first || seen=second
[ "$seen" = "none" ] || fail "or chain should stop after first success"
echo PASS
exit 0
