#!/usr/bin/env bash
# vybe-test: bash/parameter_length_expansion/param_len_special_param_question_mark
# The ${#?} expansion evaluates to the character length of the decimal exit status string.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
( exit 0 )
[ "${#?}" -eq 1 ] || fail "len of status 0: want 1, got ${#?}"

( exit 127 )
[ "${#?}" -eq 3 ] || fail "len of status 127: want 3, got ${#?}"
echo PASS
exit 0
