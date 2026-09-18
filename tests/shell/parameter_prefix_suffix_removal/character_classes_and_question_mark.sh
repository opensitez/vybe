#!/usr/bin/env bash
# vybe-test: bash/parameter_prefix_suffix_removal/character_classes_and_question_mark
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x=12ab34
[ "${x#[0-9]}" = 2ab34 ] || fail "one digit: got [${x#[0-9]}]"
[ "${x##*[0-9]}" = "" ] || fail "longest through last digit: got [${x##*[0-9]}]"
[ "${x%[[:digit:]][[:digit:]]}" = 12ab ] || fail "two-digit suffix: got [${x%[[:digit:]][[:digit:]]}]"
[ "${x#??}" = ab34 ] || fail "?? removes exactly two: got [${x#??}]"
[ "${x#[!0-9]}" = 12ab34 ] || fail "negated class must not match a digit: got [${x#[!0-9]}]"
echo PASS
exit 0
