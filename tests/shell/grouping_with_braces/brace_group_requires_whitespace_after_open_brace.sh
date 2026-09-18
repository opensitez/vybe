#!/usr/bin/env bash
# vybe-test: bash/grouping_with_braces/brace_group_requires_whitespace_after_open_brace
# '{' is a reserved word and must be followed by whitespace; '{cmd;' is treated as command name.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
eval '{echo 1; }' 2>/dev/null
st=$?
[ "$st" -ne 0 ] || fail "unspaced '{echo' should fail as command not found"
echo PASS
exit 0
