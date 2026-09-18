#!/usr/bin/env bash
# vybe-test: bash/test_builtin_forms/bracket_form_requires_closing_bracket
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
test a = a || fail "test form"
[ a = a ] || fail "bracket form"
msg=$( [ a = a 2>&1 ); st=$?
[ "$st" -eq 2 ] || fail "missing ]: want status 2 got $st"
[[ $msg == *"missing"*"]"* ]] || fail "missing ]: got [$msg]"
test a = a ] 2>/dev/null; st=$?
[ "$st" -eq 2 ] || fail "] is an ordinary argument to test: want 2 got $st"
[ "]" = "]" ] || fail "] can still be an operand"
echo PASS
exit 0
