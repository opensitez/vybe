#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/omitted_pattern_equals_question_mark
# An omitted pattern is treated as ?, which matches every character; * behaves
# the same because each character is matched individually.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x="mixed Case"
[ "${x^^?}" = "${x^^}" ] || fail "^^? differs from ^^"
[ "${x^^*}" = "${x^^}" ] || fail "^^* differs from ^^"
[ "${x,?}" = "${x,}" ] || fail ",? differs from ,"
echo PASS
exit 0
