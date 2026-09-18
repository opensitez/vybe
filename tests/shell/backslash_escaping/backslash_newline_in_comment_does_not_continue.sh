#!/usr/bin/env bash
# vybe-test: bash/backslash_escaping/backslash_newline_in_comment_does_not_continue
# A comment ends at the newline even if it ends with a backslash.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
ran=no
# this comment ends with a backslash \
ran=yes
[ "$ran" = yes ] || fail "line after backslash-terminated comment must run"
echo PASS
exit 0
