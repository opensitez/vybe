#!/usr/bin/env bash
# vybe-test: bash/integer_comparison_operators/single_bracket_requires_integer_operands
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
msg=$( [ abc -eq 0 ] 2>&1 ); st=$?
[ "$st" -eq 2 ] || fail "abc: want status 2 got $st"
[[ $msg == *"integer expected"* ]] || fail "abc: got [$msg]"
[ 1+1 -eq 2 ] 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "1+1: want status 2 got $st"
[ "" -eq 0 ] 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "empty: want status 2 got $st"
echo PASS
exit 0
