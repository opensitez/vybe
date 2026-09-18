#!/usr/bin/env bash
# vybe-test: bash/environment_inheritance/env_trailing_whitespace_preserved_in_child
# Trailing whitespace in an environment variable value is preserved verbatim when inherited by a child.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
export PADDED_STR="content   "
child_len=$( "$BASH" -c 'printf "%s\n" "${#PADDED_STR}"' )
[ "$child_len" -eq 10 ] || fail "trailing spaces trimmed across process inheritance: want length 10, got $child_len"
echo PASS
exit 0
