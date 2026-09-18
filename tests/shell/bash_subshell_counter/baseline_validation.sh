#!/usr/bin/env bash
# vybe-test: bash/bash_subshell_counter/baseline_validation
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
IDX=1
parent=$BASHPID
sub=$( (printf '%s\n' "$BASHPID") )
(( sub != parent )) || fail "subshell BASHPID should differ"
if (( IDX % 2 == 0 )); then
  nested=$( (printf '%s\n' "$BASHPID"; printf '%s\n' "$BASHPID" ) )
  first=${nested%%'
'*}
  second=${nested##*'
'}
  (( second > 0 )) || fail "nested subshell output missing"
  (( first > 0 )) || fail "nested subshell output missing"
fi
echo PASS
exit 0
