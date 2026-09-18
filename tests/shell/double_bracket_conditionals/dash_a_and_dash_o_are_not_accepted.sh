#!/usr/bin/env bash
# vybe-test: bash/double_bracket_conditionals/dash_a_and_dash_o_are_not_accepted
# Only && and || combine expressions inside [[ ]]; -a and -o are syntax errors.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( eval '[[ 1 -eq 1 -a 2 -eq 2 ]]' 2>&1 ); st=$?
[ "$st" -ne 0 ] || fail "-a must be rejected"
[[ $msg == *"syntax error in conditional expression"* ]] || fail "got [$msg]"
eval '[[ 1 -eq 1 -o 2 -eq 2 ]]' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "-o must be rejected"
echo PASS
exit 0
