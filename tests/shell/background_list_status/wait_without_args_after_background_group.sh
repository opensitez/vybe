#!/usr/bin/env bash
# vybe-test: bash/background_list_status/wait_without_args_after_background_group
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
(
  :
) &
(
  false
) &
wait
[ "$?" -eq 1 ] || fail "group has failure, wait must fail"
echo PASS
exit 0
