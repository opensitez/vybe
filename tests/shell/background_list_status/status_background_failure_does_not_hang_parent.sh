#!/usr/bin/env bash
# vybe-test: bash/background_list_status/status_background_failure_does_not_hang_parent
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
false &
p=$!
start=$(date +%s)
wait "$p"
end=$(date +%s)
[ "$?" -eq 1 ] || fail "false should fail"
[ "$((end-start))" -le 1 ] || fail "waiting failed job must return quickly"
echo PASS
exit 0
