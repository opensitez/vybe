#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/single_form_only_converts_when_first_char_matches
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="hello"
[ "${x^[a-z]}" = "Hello" ] || fail "first char matches: got [${x^[a-z]}]"
[ "${x^[e]}" = "hello" ] || fail "pattern matches only a later char, nothing changes: got [${x^[e]}]"
n="1abc"
[ "${n^^[a-z]}" = "1ABC" ] || fail "all form skips the digit: got [${n^^[a-z]}]"
[ "${n^}" = "1abc" ] || fail "digit has no case: got [${n^}]"
echo PASS
exit 0
