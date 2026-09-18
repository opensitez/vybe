#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_plus_with_slash_path_expands_to_pwd_subpath
# The ~+/subpath syntax expands to the PWD variable concatenated with the trailing subpath.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
PWD="/custom/working/directory"
res=~+/subfolder/config.json
[ "$res" = "/custom/working/directory/subfolder/config.json" ] || fail "~+/subpath failed: got [$res]"
echo PASS
exit 0
