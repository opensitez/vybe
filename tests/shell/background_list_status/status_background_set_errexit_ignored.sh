#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_background_set_errexit_ignored
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -e
false &
p=$!
set +e
wait "$p"
set -e
[ "$?" -eq 1 ] || fail "errexit in parent should not kill wait behavior"
echo PASS
exit 0
