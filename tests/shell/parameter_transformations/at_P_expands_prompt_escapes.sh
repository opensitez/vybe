#!/usr/bin/env bash
# vybe-test: bash/parameter_transformations/at_P_expands_prompt_escapes
# @P interprets the value like PS1: \n is a newline, \\ a backslash, \a BEL.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
x='a\nb'
[ "${x@P}" = $'a\nb' ] || fail "newline: got [${x@P}]"
y='a\\b'
[ "${y@P}" = 'a\b' ] || fail "backslash: got [${y@P}]"
z='\a'
[ "${z@P}" = $'\a' ] || fail "bell: got [${z@P}]"
echo PASS
exit 0
