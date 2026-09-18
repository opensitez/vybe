#!/usr/bin/env bash
# vybe-test: bash/builtin_getopts_short_options/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- -ab
OPTIND=1
getopts "ab" opt "$@"
[ "$?" -eq 0 ] || fail "first short opt failed"
[ "$opt" = "a" ] || fail "first short opt mismatch"
getopts "ab" opt "$@"
[ "$?" -eq 0 ] || fail "second short opt failed"
[ "$opt" = "b" ] || fail "second short opt mismatch"
[ "$OPTIND" -eq 2 ] || fail "OPTIND should stay for grouped options"
echo PASS
exit 0
