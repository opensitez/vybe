#!/usr/bin/env bash
# vybe-test: bash/background_list_execution/background_disown_job_still_can_be_waited
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( : ) &
p=$!
disown "$p"
wait "$p"
st=$?
[ "$st" -eq 0 ] || fail "disowned pid must still be waitable"
echo PASS
exit 0
