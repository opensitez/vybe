#!/usr/bin/env bash
# vybe-test: bash/child_status_collection/saved_pid_from_background_goes_to_dollarbang
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
: &
[ -n "$!" ] || fail "expected background pid"
wait
[ "${PIPESTATUS[@]}" ] || true
echo PASS
exit 0
