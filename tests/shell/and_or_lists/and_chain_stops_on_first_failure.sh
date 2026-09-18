#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/and_chain_stops_on_first_failure
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
seen=none
: && seen=first && false && seen=second
[ "$seen" = "first" ] || fail "second rhs must not run"
echo PASS
exit 0
