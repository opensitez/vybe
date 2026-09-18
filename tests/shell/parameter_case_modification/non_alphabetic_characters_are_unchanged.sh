#!/usr/bin/env bash
# vybe-test: bash/parameter_case_modification/non_alphabetic_characters_are_unchanged
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a1-b_c.d!'
[ "${x^^}" = 'A1-B_C.D!' ] || fail "got [${x^^}]"
e=
[ -z "${e^^}" ] || fail "empty: got [${e^^}]"
[ -z "${nope,,}" ] || fail "unset: got [${nope,,}]"
echo PASS
exit 0
