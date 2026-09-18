#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/pattern_selects_which_characters_change
# The optional pattern must match a single character; only matching
# characters are converted.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="hello world"
[ "${x^^[aeiou]}" = "hEllO wOrld" ] || fail "vowels: got [${x^^[aeiou]}]"
y="HELLO"
[ "${y,,[HL]}" = "hEllO" ] || fail "class: got [${y,,[HL]}]"
[ "${x^^[!a-z]}" = "hello world" ] || fail "no letters match [!a-z]: got [${x^^[!a-z]}]"
echo PASS
exit 0
