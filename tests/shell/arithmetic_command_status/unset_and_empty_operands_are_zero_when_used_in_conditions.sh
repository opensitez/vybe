#!/usr/bin/env bash
# vybe-test: bash/arithmetic_command_status/unset_and_empty_operands_are_zero_when_used_in_conditions
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
unset z
zstatus=unset
(( z )); zstatus=$?
[ "$zstatus" -eq 1 ] || fail "unset name in (( )) should behave as 0"
empty=''
(( empty )); estatus=$?
[ "$estatus" -eq 1 ] || fail "empty name in (( )) should behave as 0"
echo PASS
exit 0
