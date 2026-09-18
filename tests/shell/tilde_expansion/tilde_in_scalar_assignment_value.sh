#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_in_scalar_assignment_value
# In variable assignment syntax (var=~/dir), the tilde immediately following the '=' undergoes tilde expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
target=~/downloads
[ "$target" = "$HOME/downloads" ] || fail "assignment tilde failed: want [$HOME/downloads], got [$target]"
echo PASS
exit 0
