#!/usr/bin/env bash
# vybe-test: bash/and_or_lists/newlines_do_not_change_or_list_evaluation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
log=""
: \
  || \
  log=split
[ "$log" = "split" ] || fail "newline split should not alter semantics"
echo PASS
exit 0
