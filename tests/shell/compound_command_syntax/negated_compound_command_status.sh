#!/usr/bin/env bash
# vybe-test: bash/compound_command_syntax/negated_compound_command_status
# The '!' reserved word negates the overall exit status of a compound command.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
! { (exit 1); }
st1=$?
[ "$st1" -eq 0 ] || fail "! { exit 1; }: want 0, got $st1"

! { (exit 0); }
st2=$?
[ "$st2" -eq 1 ] || fail "! { exit 0; }: want 1, got $st2"
echo PASS
exit 0
