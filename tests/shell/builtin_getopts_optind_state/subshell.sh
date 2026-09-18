#!/usr/bin/env bash
# vybe-test: bash/builtin_getopts_optind_state/subshell
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
set -- -a "ok" -b "x"
OPTIND=1
while getopts ":a:b:" opt "$@"; do
  :
done
[ "$OPTIND" -eq 5 ] || fail "OPTIND not advanced to expected position"
echo PASS
exit 0
