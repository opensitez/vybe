#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_with_slash_path_expands_to_home_subpath
# The ~/subpath syntax expands to $HOME concatenated with the trailing subpath.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
res=~/documents/projects
[ "$res" = "$HOME/documents/projects" ] || fail "tilde subpath failed: want [$HOME/documents/projects], got [$res]"
echo PASS
exit 0
