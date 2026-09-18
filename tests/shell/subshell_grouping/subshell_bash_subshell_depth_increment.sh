#!/usr/bin/env bash
# vybe-test: bash/subshell_grouping/subshell_bash_subshell_depth_increment
# The BASH_SUBSHELL special variable increments by 1 for each nested subshell level.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
top_depth=$BASH_SUBSHELL
[ "$top_depth" -eq 0 ] || fail "top-level depth: want 0, got $top_depth"

( exit "$BASH_SUBSHELL" )
d1=$?
[ "$d1" -eq 1 ] || fail "level 1 subshell depth: want 1, got $d1"

( ( exit "$BASH_SUBSHELL" ) )
d2=$?
[ "$d2" -eq 2 ] || fail "level 2 subshell depth: want 2, got $d2"

( ( ( exit "$BASH_SUBSHELL" ) ) )
d3=$?
[ "$d3" -eq 3 ] || fail "level 3 subshell depth: want 3, got $d3"
echo PASS
exit 0
