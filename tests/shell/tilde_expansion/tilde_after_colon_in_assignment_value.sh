#!/usr/bin/env bash
# vybe-test: bash/tilde_expansion/tilde_after_colon_in_assignment_value
# In assignment statements, each unquoted tilde immediately following a colon ':' undergoes tilde expansion.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
[ -n "$HOME" ] || fail "HOME is not set"
search_path=~/bin:/usr/local/bin:~/scripts
expected="$HOME/bin:/usr/local/bin:$HOME/scripts"
[ "$search_path" = "$expected" ] || fail "colon tildes failed: want [$expected], got [$search_path]"
echo PASS
exit 0
