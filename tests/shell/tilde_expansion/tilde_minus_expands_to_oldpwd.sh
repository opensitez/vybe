#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_minus_expands_to_oldpwd
# An unquoted '~-' expands to the value of the OLDPWD variable (previous working directory).
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
OLDPWD="/custom/previous/directory"
res=~-
[ "$res" = "/custom/previous/directory" ] || fail "~- expansion failed: want '/custom/previous/directory', got [$res]"
echo PASS
exit 0
