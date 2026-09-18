#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_with_nonexistent_user_stays_literal
# If the username in ~username is not a valid login name, the string remains untouched as literal text.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
res=~nonexistent_user_identifier_12345
[ "$res" = "~nonexistent_user_identifier_12345" ] || fail "nonexistent user tilde should stay literal: got [$res]"
echo PASS
exit 0
