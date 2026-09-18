#!/usr/bin/env bash
# vybe-test: bash/special_parameters/special_param_length_expansion_on_question_mark
# The ${#?} expansion calculates the character length of the exit status in decimal.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( exit 0 )
[ "${#?}" -eq 1 ] || fail "length of status 0: want 1, got ${#?}"

( exit 127 )
[ "${#?}" -eq 3 ] || fail "length of status 127: want 3, got ${#?}"
echo PASS
exit 0
