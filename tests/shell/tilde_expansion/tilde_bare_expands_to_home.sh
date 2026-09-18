#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_bare_expands_to_home
# A bare unquoted tilde '~' expands to the user's home directory matching $HOME.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME environment variable is not set"
res=~
[ "$res" = "$HOME" ] || fail "bare tilde failed: want [$HOME], got [$res]"
echo PASS
exit 0
