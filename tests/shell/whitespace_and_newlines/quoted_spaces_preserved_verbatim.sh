#!/usr/bin/env bash
# vybe-test: bash/whitespace_and_newlines/quoted_spaces_preserved_verbatim
# Spaces inside double or single quotes are not collapsed and remain exact.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
s1="  multiple   spaces  "
s2='  single   quoted  '
[ "${#s1}" -eq 21 ] || fail "s1 length: want 21, got ${#s1}"
[ "${#s2}" -eq 19 ] || fail "s2 length: want 19, got ${#s2}"
echo PASS
exit 0
