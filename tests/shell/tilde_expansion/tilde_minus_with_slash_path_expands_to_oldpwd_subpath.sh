#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_minus_with_slash_path_expands_to_oldpwd_subpath
# The ~-/subpath syntax expands to the OLDPWD variable concatenated with the trailing subpath.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
OLDPWD="/custom/previous/directory"
res=~-/history.log
[ "$res" = "/custom/previous/directory/history.log" ] || fail "~-/subpath failed: got [$res]"
echo PASS
exit 0
